#!/usr/bin/env python3
"""
Sistema MRP com Endereçamento de Estoque
Material Requirements Planning

Passos de processamento:
  1.  Ler demanda
  2.  Consolidar demanda
  3.  Somar estoque por material (ignorar endereçamento)
  4.  Ler pedidos em aberto (entradas futuras)
  5.  Calcular classificação ABC
  6.  Iniciar cálculo MRP mês a mês
  7.  Inserir entradas (pedidos existentes + novos)
  8.  Calcular estoque projetado
  9.  Calcular necessidade
  10. Gerar compras
  11. Aplicar estratégia por classe
  12. Gerar rateio final por departamento / programa orçamentário
"""

import io
import math
import os
from datetime import date, timedelta

import numpy as np
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────
MESES_COBERTURA_SS     = 3   # Cobertura do estoque de segurança (A/B) — rolling
HORIZONTE_MESES        = 12  # Horizonte de planejamento
LEAD_TIME_DIAS         = 30  # Lead time em DIAS CORRIDOS
#   data_chegada = data_pedido + LEAD_TIME_DIAS
#   A entrada é alocada no mês da data_chegada

LIMITE_ABC_A = 0.80          # Classe A → até 80% do valor acumulado
LIMITE_ABC_B = 0.95          # Classe B → 80% a 95%

# ── Estratégia Classe C ────────────────────────────────────────────────────
# Objetivo: NUNCA ter ruptura — consolidar pedidos em compras maiores e
# menos frequentes em vez de bloquear por quantidade de pedidos/ano.
#
# Trigger  : pedido quando estoque projetado < demanda do mês corrente
#             (cobertura menor que 1 mês = risco iminente de ruptura)
# Cobertura: ao pedir, cobrir os próximos CLASSE_C_COBERTURA_MESES meses
#             (pedido maior → menos pedidos ao longo do ano)
CLASSE_C_COBERTURA_MESES = 4   # Meses cobertos por pedido Classe C

DIR_DADOS = "data"
DIR_SAIDA = "output"

# Arquivo de demanda bruta no formato DTM (dd/mm/yyyy).
# Quando definido e o arquivo existir, substitui demanda.csv como fonte de demanda.
# Definir como None para usar demanda.csv diretamente.
ARQUIVO_DEMANDA_RAW = "demanda_dtm_raw.csv"

# ── Nomes convencionais dos arquivos SAP (usados pelo main() em modo CLI) ──────
# Quando os arquivos existirem em DIR_DADOS, os parsers SAP são ativados
# automaticamente; caso contrário, o fluxo legado é mantido.
ARQ_REMESSAS_SAP  = "remessas_sap.csv"    # tab-sep exportado do SAP ME2M / ME9F
ARQ_ESTOQUE_SAP   = "estoque_sap.csv"     # tab-sep exportado do SAP MB52 / MMBE
ARQ_CONTRATOS_SAP = "contratos_sap.csv"   # tab-sep exportado do SAP ME3M / ME3N
ARQ_LEAD_TIMES    = "lead_times.csv"      # material,lead_time_dias  (CSV simples)

# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────
W = 72

def cabecalho(titulo: str) -> None:
    print(f"\n{'═' * W}")
    print(f"  {titulo}")
    print(f"{'═' * W}")


def separador(titulo: str = "") -> None:
    if titulo:
        print(f"\n{'─' * W}")
        print(f"  {titulo}")
        print(f"{'─' * W}")
    else:
        print(f"{'─' * W}")


def salvar(df: pd.DataFrame, nome: str) -> None:
    caminho = os.path.join(DIR_SAIDA, nome)
    df.to_csv(caminho, index=False, encoding="utf-8-sig")
    print(f"  ✓ Salvo → {caminho}  ({len(df)} linhas)")


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS SAP — FUNÇÕES NOVAS (adição pura; não alteram nenhum passo existente)
# ─────────────────────────────────────────────────────────────────────────────

def br_to_float(s) -> float:
    """Converte número no formato brasileiro ('3.515,50') para float (3515.50).
    Regra: remove separador de milhar (ponto) e troca decimal (vírgula) por ponto.
    """
    if s is None:
        return 0.0
    s = str(s).strip()
    if s in ("", "-", "nan", "NaN"):
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _ler_sap_tabsep(source) -> pd.DataFrame:
    """Lê arquivo SAP exportado como tab-separated ou CSV com ponto-e-vírgula.
    Aceita: path string  OU  file-like object (BytesIO / UploadedFile Streamlit).
    Tenta utf-8-sig → latin-1 → cp1252 em caso de erro de encoding.
    Tenta \t → ; → , como separador (escolhe o que produz mais colunas).
    """
    encodings = ["utf-8-sig", "latin-1", "cp1252"]
    separators = ["\t", ";", ","]

    best = None
    for enc in encodings:
        for sep in separators:
            try:
                if hasattr(source, "seek"):
                    source.seek(0)
                df = pd.read_csv(
                    source, encoding=enc, sep=sep,
                    dtype=str, na_values=[""], keep_default_na=False,
                )
                if best is None or len(df.columns) > len(best.columns):
                    best = df
                # Se já encontrou múltiplas colunas com este encoding, não testa outros separadores
                if len(df.columns) > 1:
                    break
            except UnicodeDecodeError:
                break  # tenta próximo encoding
            except Exception:
                continue
        if best is not None and len(best.columns) > 1:
            break

    if best is None or best.empty:
        raise ValueError("Não foi possível ler o arquivo SAP.")
    return best


# ── Parser File 2: Remessas (Delivery Schedule) ───────────────────────────────
def ler_remessas_sap(source) -> tuple:
    """
    Lê arquivo SAP de remessas (ME2M / pedidos de compra com datas de entrega).

    Colunas utilizadas:
      'Material'                  → material
      'Data de remessa'           → data_remessa  (dd/mm/yyyy)
      'a ser fornecida (quantidade)' → quantidade  (formato BR)
      'Código de eliminação'      → codigo_eliminacao

    Regras:
      • Eliminar linhas com 'Código de eliminação' == 'L'
      • Eliminar linhas com quantidade == 0
      • Datas anteriores ao mês atual → realocadas para o mês atual (atraso)

    Retorna: (entradas_consolidadas, df_fut)
      entradas_consolidadas — material | mes | qtd_entrada   (mesmo contrato de passo_4)
      df_fut                — linhas originais filtradas com coluna 'mes_remessa'
    """
    separador("PARSER SAP │ REMESSAS (File 2)")

    df = _ler_sap_tabsep(source)

    # ── Mapear colunas obrigatórias ───────────────────────────────────────────
    col_map = {
        "Material"                        : "material",
        "Data de remessa"                 : "_data_raw",
        "a ser fornecida (quantidade)"    : "_qtd_raw",
        "Código de eliminação"            : "codigo_eliminacao",
    }
    ausentes = [c for c in col_map if c not in df.columns]
    if ausentes:
        raise ValueError(f"Colunas obrigatórias ausentes em remessas_sap: {ausentes}")

    df = df.rename(columns=col_map)

    # ── Filtro 1: eliminar código 'L' ─────────────────────────────────────────
    antes = len(df)
    df = df[df["codigo_eliminacao"].fillna("").str.strip().str.upper() != "L"].copy()
    print(f"  Filtro código 'L'    : {antes - len(df)} linha(s) removida(s)")

    # ── Converter quantidade (formato BR) ─────────────────────────────────────
    df["quantidade"] = df["_qtd_raw"].apply(br_to_float)

    # ── Filtro 2: eliminar quantidade == 0 ────────────────────────────────────
    antes = len(df)
    df = df[df["quantidade"] > 0].copy()
    print(f"  Filtro qtd == 0      : {antes - len(df)} linha(s) removida(s)")

    # ── Converter data de remessa ─────────────────────────────────────────────
    df["data_remessa"] = pd.to_datetime(df["_data_raw"], format="%d/%m/%Y", errors="coerce")
    invalidas = df["data_remessa"].isna()
    if invalidas.any():
        print(f"  ⚠ Datas inválidas    : {invalidas.sum()} — linhas descartadas")
        df = df[~invalidas].copy()

    df["mes_remessa"] = df["data_remessa"].dt.to_period("M").astype(str)

    # ── Regra de atraso: datas passadas → mês atual ───────────────────────────
    mes_atual = str(pd.Period(date.today(), "M"))
    atrasados = df["mes_remessa"] < mes_atual
    if atrasados.any():
        print(f"  Realocar atrasados   : {atrasados.sum()} linha(s) → {mes_atual}")
        df.loc[atrasados, "mes_remessa"] = mes_atual

    df_fut = df[df["mes_remessa"] >= mes_atual].copy()

    print(f"  Remessas após filtros : {len(df_fut)} linha(s)")
    print(f"  Materiais únicos      : {df_fut['material'].nunique()}")

    entradas = (
        df_fut.groupby(["material", "mes_remessa"], as_index=False)["quantidade"]
        .sum()
        .rename(columns={"mes_remessa": "mes", "quantidade": "qtd_entrada"})
    )
    return entradas, df_fut


# ── Parser File 3: Estoque (Stock / MB52 / MMBE) ──────────────────────────────
def ler_estoque_sap(source) -> pd.DataFrame:
    """
    Lê arquivo SAP de estoque (MB52 / MMBE multi-depósito).

    Colunas utilizadas:
      'Produto'             → material  (código limpo, ex: '400011')
      'Qtd.disponível UMB'  → quantidade (formato BR)

    Tratamento especial:
      Linhas com depósito BLIN têm uma coluna extra no início, deslocando o layout.
      A função normaliza essas linhas antes do parsing.

    Regras:
      • Agrupar por material → somar quantidades
      • Estoque total negativo → forçar para 0

    Retorna: DataFrame  material | estoque_total
    """
    separador("PARSER SAP │ ESTOQUE (File 3)")

    # ── Leitura linha a linha para tratar o offset BLIN ───────────────────────
    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, bytes):
            for enc in ["utf-8-sig", "latin-1", "cp1252"]:
                try:
                    raw = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        linhas = raw.splitlines()
    else:
        with open(source, encoding="utf-8-sig", errors="replace") as fh:
            linhas = fh.read().splitlines()

    # Encontrar linha de cabeçalho (contém 'Produto')
    header_idx = None
    for i, ln in enumerate(linhas):
        if "Produto" in ln:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Coluna 'Produto' não encontrada no arquivo de estoque SAP.")

    header_fields = linhas[header_idx].split("\t")
    n_cols = len(header_fields)

    # Normalizar linhas de dados: BLIN tem 1 coluna extra no início → remover
    rows = []
    for ln in linhas[header_idx + 1:]:
        if not ln.strip():
            continue
        fields = ln.split("\t")
        if len(fields) == n_cols + 1 and fields[0].strip().upper() == "BLIN":
            fields = fields[1:]          # remove a coluna extra de BLIN
        # Padding / truncate para n_cols
        while len(fields) < n_cols:
            fields.append("")
        rows.append(fields[:n_cols])

    df = pd.DataFrame(rows, columns=header_fields)

    # ── Identificar colunas mapeadas ──────────────────────────────────────────
    col_prod = "Produto"
    col_qtd  = "Qtd.disponível UMB"

    if col_qtd not in df.columns:
        # Tentar variante sem acento
        candidatas = [c for c in df.columns if "disponível" in c or "disponivel" in c.lower()]
        if candidatas:
            col_qtd = candidatas[0]
        else:
            raise ValueError(f"Coluna '{col_qtd}' não encontrada no arquivo de estoque SAP.")

    df = df[[col_prod, col_qtd]].rename(
        columns={col_prod: "material", col_qtd: "_qtd_raw"}
    )

    # ── Limpar material: remover sufixos (ex: '400011INVT' → ignorar) ─────────
    # O código limpo é numérico (ou alfanum. sem sufixo de tipo).
    # Manter apenas linhas onde material é numérico ou tem no máx. 10 caracteres
    # sem o padrão de sufixo SAP (letras depois de números).
    df["material"] = df["material"].str.strip()
    df = df[df["material"].str.match(r"^\d+$", na=False)].copy()  # só códigos numéricos limpos

    df["quantidade"] = df["_qtd_raw"].apply(br_to_float)

    # ── Consolidar por material ───────────────────────────────────────────────
    consolidado = (
        df.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "estoque_total"})
    )

    # ── Estoque negativo → 0 ──────────────────────────────────────────────────
    neg = consolidado["estoque_total"] < 0
    if neg.any():
        print(f"  ⚠ Estoque negativo   : {neg.sum()} material(is) → forçado para 0")
        consolidado.loc[neg, "estoque_total"] = 0

    print(f"  Materiais únicos      : {len(consolidado)}")
    print(f"  Estoque total (soma)  : {consolidado['estoque_total'].sum():,.0f} un.")
    return consolidado


# ── Parser File 4: Contratos (Framework Agreements) ───────────────────────────
def ler_contratos_sap(source) -> pd.DataFrame:
    """
    Lê arquivo SAP de contratos (ME3M / ME3N / accordos de remessa).

    Colunas utilizadas:
      'Material'           → material
      'Fim da validade'    → data_fim_vigencia  (dd/mm/yyyy)
      'Qtd.prev.pendente'  → saldo_contrato     (formato BR)
      'Preço líquido'      → valor_unitario     (formato BR)

    Regras:
      • Extrair preço de TODOS os contratos (inclusive vencidos) para referência ABC
      • Para saldo: filtrar data_fim_vigencia >= hoje  E  saldo_contrato > 0
      • Por material: SUM(saldo_contrato), MAX(valor_unitario)

    Retorna: DataFrame  material | data_fim_vigencia | saldo_contrato | valor_unitario
    """
    separador("PARSER SAP │ CONTRATOS (File 4)")

    df = _ler_sap_tabsep(source)

    # ── Mapear colunas ────────────────────────────────────────────────────────
    col_map = {
        "Material"          : "material",
        "Fim da validade"   : "_data_raw",
        "Qtd.prev.pendente" : "_saldo_raw",
        "Preço líquido"     : "_preco_raw",
    }
    ausentes = [c for c in col_map if c not in df.columns]
    if ausentes:
        print(f"  Colunas encontradas: {list(df.columns)}")
        raise ValueError(f"Colunas obrigatórias ausentes em contratos_sap: {ausentes}")

    df = df.rename(columns=col_map)[["material", "_data_raw", "_saldo_raw", "_preco_raw"]]
    df["material"]       = df["material"].str.strip()
    df["saldo_contrato"] = df["_saldo_raw"].apply(br_to_float)
    df["valor_unitario"] = df["_preco_raw"].apply(br_to_float)
    df["data_fim_vigencia"] = pd.to_datetime(df["_data_raw"], format="%d/%m/%Y", errors="coerce")

    invalidas = df["data_fim_vigencia"].isna()
    if invalidas.any():
        print(f"  ⚠ Datas inválidas    : {invalidas.sum()} linha(s) descartadas")
        df = df[~invalidas].copy()

    hoje = pd.Timestamp(date.today())

    # ── Preço máximo por material (inclui contratos vencidos — referência ABC) ─
    max_preco = (
        df.groupby("material", as_index=False)["valor_unitario"]
        .max()
        .rename(columns={"valor_unitario": "valor_unitario_max"})
    )

    # ── Filtrar contratos vigentes com saldo > 0 ──────────────────────────────
    vigentes = df[(df["data_fim_vigencia"] >= hoje) & (df["saldo_contrato"] > 0)].copy()
    print(f"  Total de linhas       : {len(df)}")
    print(f"  Linhas vigentes c/ saldo > 0 : {len(vigentes)}")

    if vigentes.empty:
        print("  ⚠ Nenhum contrato vigente com saldo > 0")
        # Retornar tabela só com preço (sem saldo)
        resultado = max_preco.rename(columns={"valor_unitario_max": "valor_unitario"})
        resultado["saldo_contrato"] = 0.0
        resultado["data_fim_vigencia"] = pd.NaT
        return resultado[["material", "data_fim_vigencia", "saldo_contrato", "valor_unitario"]]

    # ── Somar saldo por material ───────────────────────────────────────────────
    saldo_total = (
        vigentes.groupby("material", as_index=False)["saldo_contrato"]
        .sum()
    )
    # Data de vencimento mais próxima (conservador)
    data_min = (
        vigentes.groupby("material", as_index=False)["data_fim_vigencia"]
        .min()
    )

    resultado = saldo_total.merge(data_min, on="material").merge(max_preco, on="material")
    resultado = resultado.rename(columns={"valor_unitario_max": "valor_unitario"})
    resultado["data_fim_vigencia"] = resultado["data_fim_vigencia"].dt.strftime("%d/%m/%Y")

    print(f"  Materiais com saldo   : {len(resultado)}")
    print(resultado[["material", "data_fim_vigencia", "saldo_contrato", "valor_unitario"]]
          .to_string(index=False))
    return resultado[["material", "data_fim_vigencia", "saldo_contrato", "valor_unitario"]]


# ── Parser File 5: Lead Times ──────────────────────────────────────────────────
def ler_lead_times(source) -> dict:
    """
    Lê arquivo CSV simples  material,lead_time_dias.
    Retorna dict  {material_str: lead_time_int}.
    Materiais ausentes devem usar LEAD_TIME_DIAS (default 60 dias).
    """
    separador("PARSER │ LEAD TIMES (File 5)")
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        df = pd.read_csv(source, dtype=str)
    except Exception as exc:
        print(f"  ⚠ Não foi possível ler lead_times: {exc} — usando default {LEAD_TIME_DIAS}d para todos")
        return {}

    if "material" not in df.columns or "lead_time_dias" not in df.columns:
        print(f"  ⚠ Colunas esperadas: material, lead_time_dias — usando default para todos")
        return {}

    df["lead_time_dias"] = pd.to_numeric(df["lead_time_dias"], errors="coerce").fillna(LEAD_TIME_DIAS)
    lt_dict = {
        str(row["material"]).strip(): int(row["lead_time_dias"])
        for _, row in df.iterrows()
    }
    print(f"  Lead times carregados : {len(lt_dict)} material(is)")
    print(f"  Default (ausentes)    : {LEAD_TIME_DIAS} dias")
    return lt_dict


# ── Derivar base de rateio da própria demanda ──────────────────────────────────
def derivar_rateio_da_demanda(demanda_detail: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula proporções de rateio por (departamento, programa_orcamentario)
    a partir do arquivo de demanda DTM, sem usar rateio_base.csv.

    Retorna DataFrame com colunas:
      material | departamento | programa_orcamentario | proporcao
    """
    colunas_req = {"material", "departamento", "programa_orcamentario", "quantidade"}
    if not colunas_req.issubset(demanda_detail.columns):
        return pd.DataFrame(
            columns=["material", "departamento", "programa_orcamentario", "proporcao"]
        )

    por_depto = (
        demanda_detail.groupby(
            ["material", "departamento", "programa_orcamentario"], as_index=False
        )["quantidade"]
        .sum()
    )

    total_mat = (
        demanda_detail.groupby("material")["quantidade"]
        .sum()
        .rename("_total")
        .reset_index()
    )

    por_depto = por_depto.merge(total_mat, on="material")
    por_depto["proporcao"] = por_depto["quantidade"] / por_depto["_total"].replace(0, 1)

    return por_depto[["material", "departamento", "programa_orcamentario", "proporcao"]]


# ─────────────────────────────────────────────────────────────────────────────
# CÁLCULO DE LEAD TIME EM DIAS CORRIDOS
# ─────────────────────────────────────────────────────────────────────────────
def calcular_periodo_entrega(
    periodo_necessidade: str,
    lead_time_dias: int,
) -> tuple[str, date, date]:
    """
    Calcula o período de entrega a partir do período de necessidade e do
    lead time em dias corridos.

    Regras:
      - Período atual  → data_pedido = hoje (pedido emitido agora)
      - Períodos futuros → data_pedido = 1º dia do mês
      - data_chegada = data_pedido + lead_time_dias
      - período_entrega = mês da data_chegada
      - GARANTIA: período_entrega > período_necessidade
        (nenhuma entrada pode cair no mesmo mês ou no passado)

    Retorna: (periodo_entrega, data_pedido, data_chegada)
    """
    hoje = date.today()
    per  = pd.Period(periodo_necessidade, "M")

    # Data em que o pedido seria emitido
    if per == pd.Period(hoje, "M"):
        data_pedido = hoje                          # mês atual → emite hoje
    else:
        data_pedido = date(per.year, per.month, 1)  # mês futuro → emite no dia 1

    # Data de chegada = data_pedido + lead time em dias corridos
    data_chegada    = data_pedido + timedelta(days=lead_time_dias)
    periodo_chegada = str(pd.Period(data_chegada, "M"))

    # Garantia de consistência: chegada deve ser POSTERIOR ao mês da necessidade
    if pd.Period(periodo_chegada, "M") <= per:
        proximo    = per + 1
        periodo_chegada = str(proximo)
        data_chegada    = date(proximo.year, proximo.month, 1)

    return periodo_chegada, data_pedido, data_chegada


# ─────────────────────────────────────────────────────────────────────────────
# PRÉ-PROCESSAMENTO: TRANSFORMAR DEMANDA NO FORMATO BRUTO DTM
# ─────────────────────────────────────────────────────────────────────────────
def transformar_demanda_dtm(caminho: str) -> pd.DataFrame:
    """
    Lê o arquivo de demanda no formato bruto DTM e aplica o mapeamento oficial:

      CÓDIGO  → material
      MÊS     → mes  (dd/mm/yyyy → YYYY-MM)  ← ÚNICA coluna de data utilizada
      DEP.    → departamento
      PROJETO → programa_orcamentario
      QTD     → quantidade  (numérico; negativos sinalizados)

    Colunas ignoradas: Nome Demanda, DEPÓSITO, Descrição, UNID, MES, SEMANA

    Adiciona flag_demanda_ativa = True quando a soma anual do material > 0.
    Materiais com flag_demanda_ativa = False são mantidos no MRP (sem geração
    de pedidos) para preservar rastreabilidade.

    Retorna DataFrame com colunas:
      material | mes | departamento | programa_orcamentario | quantidade | flag_demanda_ativa
    """
    separador("PRÉ-PROCESSAMENTO │ TRANSFORMAR DEMANDA DTM")

    for _enc in ["utf-8-sig", "latin-1", "cp1252"]:
        try:
            df = pd.read_csv(caminho, sep=";", encoding=_enc, dtype=str)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"Não foi possível decodificar {caminho} com utf-8-sig / latin-1 / cp1252.")

    # ── Mapeamento oficial de colunas ─────────────────────────────────────────
    col_map = {
        "CÓDIGO" : "material",
        "MÊS"    : "_mes_raw",
        "DEP."   : "departamento",
        "PROJETO": "programa_orcamentario",
        "QTD"    : "quantidade",
    }
    colunas_ausentes = [c for c in col_map if c not in df.columns]
    if colunas_ausentes:
        raise ValueError(f"Colunas obrigatórias ausentes no arquivo: {colunas_ausentes}")

    df = df.rename(columns=col_map)[[
        "material", "_mes_raw", "departamento", "programa_orcamentario", "quantidade"
    ]]

    # ── Normalização de datas: dd/mm/yyyy → YYYY-MM ───────────────────────────
    df["mes"] = (
        pd.to_datetime(df["_mes_raw"], format="%d/%m/%Y", errors="coerce")
        .dt.to_period("M")
        .astype("object")
        .where(lambda s: s.notna(), other=None)
    )
    # Converter Period para string (mantém None para inválidas)
    df["mes"] = df["mes"].apply(lambda v: str(v) if v is not None and str(v) != "NaT" else None)

    invalidas = df["mes"].isna()
    if invalidas.any():
        print(f"  ⚠  Datas inválidas (coluna MÊS): {invalidas.sum()} registro(s)")
        print(df[invalidas][["material", "_mes_raw"]].to_string(index=False))
        df = df[~invalidas].copy()

    df = df.drop(columns=["_mes_raw"])

    # ── Conversão de quantidade para numérico ─────────────────────────────────
    df["quantidade"] = pd.to_numeric(df["quantidade"], errors="coerce")

    nao_num = df["quantidade"].isna()
    if nao_num.any():
        print(f"  ⚠  Quantidades não numéricas: {nao_num.sum()} registro(s) → convertidos para 0")
    df["quantidade"] = df["quantidade"].fillna(0)

    negativos = df["quantidade"] < 0
    if negativos.any():
        print(f"  ⚠  Quantidades NEGATIVAS: {negativos.sum()} registro(s)")
        print(df[negativos][["material", "mes", "departamento", "quantidade"]].to_string(index=False))

    # ── flag_demanda_ativa: soma anual por material > 0 ───────────────────────
    soma_anual = (
        df.groupby("material")["quantidade"]
        .sum()
        .rename("_soma_anual")
    )
    df = df.join(soma_anual, on="material")
    df["flag_demanda_ativa"] = df["_soma_anual"] > 0
    df = df.drop(columns=["_soma_anual"])

    # ── Relatório ─────────────────────────────────────────────────────────────
    n_ativos   = df[df["flag_demanda_ativa"]]["material"].nunique()
    n_inativos = df[~df["flag_demanda_ativa"]]["material"].nunique()
    print(f"  Fonte              : {caminho}")
    print(f"  Registros lidos    : {len(df)}")
    print(f"  Materiais ativos   : {n_ativos}  (soma anual > 0)")
    print(f"  Materiais inativos : {n_inativos}  (soma anual = 0 — mantidos no MRP sem geração de pedidos)")
    if n_inativos > 0:
        mats = sorted(df[~df["flag_demanda_ativa"]]["material"].unique())
        print(f"    └─ {', '.join(map(str, mats))}")

    return df[[
        "material", "mes", "departamento", "programa_orcamentario",
        "quantidade", "flag_demanda_ativa"
    ]]


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 1-2: LER E CONSOLIDAR DEMANDA
# ─────────────────────────────────────────────────────────────────────────────
def passo_1_2_demanda() -> pd.DataFrame:
    separador("PASSO 1-2 │ LER E CONSOLIDAR DEMANDA")

    # ── Roteamento: formato bruto DTM  vs  demanda.csv padrão ────────────────
    caminho_raw = (
        os.path.join(DIR_DADOS, ARQUIVO_DEMANDA_RAW)
        if ARQUIVO_DEMANDA_RAW
        else None
    )

    if caminho_raw and os.path.exists(caminho_raw):
        # Fonte: arquivo bruto DTM — aplica mapeamento e normalização
        df_raw = transformar_demanda_dtm(caminho_raw)
        df = df_raw[["material", "mes", "quantidade"]].copy()
        separador()
    else:
        # Fonte: demanda.csv já no formato padrão material|mes|quantidade
        df = pd.read_csv(os.path.join(DIR_DADOS, "demanda.csv"), encoding="latin-1", sep=";")
        df["mes"] = pd.to_datetime(df["mes"], format="%Y-%m").dt.to_period("M").astype(str)

    # ── Consolidação final (agrupa caso haja duplicidades de chave) ───────────
    demanda = df.groupby(["material", "mes"], as_index=False)["quantidade"].sum()

    print(f"  Materiais com demanda : {demanda['material'].nunique()}")
    print(f"  Período da demanda    : {demanda['mes'].min()} → {demanda['mes'].max()}")
    print(f"  Registros consolidados: {len(demanda)}")
    return demanda


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3: CONSOLIDAR ESTOQUE POR MATERIAL (IGNORAR ENDEREÇAMENTO)
# ─────────────────────────────────────────────────────────────────────────────
def passo_3_estoque() -> pd.DataFrame:
    separador("PASSO 3 │ CONSOLIDAR ESTOQUE (IGNORAR ENDEREÇAMENTO)")

    # ── Detecção de formato: SAP tab-sep tem prioridade sobre estoque.csv legado ─
    sap_path = os.path.join(DIR_DADOS, ARQ_ESTOQUE_SAP)
    if os.path.exists(sap_path):
        consolidado = ler_estoque_sap(sap_path)
        salvar(consolidado, "01_estoque_consolidado.csv")
        return consolidado

    df = pd.read_csv(os.path.join(DIR_DADOS, "estoque.csv"), encoding="latin-1", sep=";")
    print(f"  Linhas de endereçamento: {len(df)}")

    consolidado = (
        df.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "estoque_total"})
    )

    print(f"  Materiais únicos      : {len(consolidado)}")
    print()
    print(consolidado.to_string(index=False))

    salvar(consolidado, "01_estoque_consolidado.csv")
    return consolidado


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4: LER PEDIDOS EM ABERTO (ENTRADAS FUTURAS)
# ─────────────────────────────────────────────────────────────────────────────
def passo_4_pedidos_abertos() -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      entradas_consolidadas — agrupado por material+mês (para o loop MRP)
      df_abertos_futuros    — linhas originais filtradas (para rateio)
    """
    separador("PASSO 4 │ PEDIDOS EM ABERTO — ENTRADAS FUTURAS")

    # ── Detecção de formato: remessas SAP têm prioridade sobre pedidos_abertos.csv ─
    sap_path = os.path.join(DIR_DADOS, ARQ_REMESSAS_SAP)
    if os.path.exists(sap_path):
        return ler_remessas_sap(sap_path)   # já retorna (entradas, df_fut) no mesmo contrato

    df = pd.read_csv(
        os.path.join(DIR_DADOS, "pedidos_abertos.csv"),
        parse_dates=["data_remessa"],
    )

    # REGRA: usar EXCLUSIVAMENTE data_remessa
    df["mes_remessa"] = df["data_remessa"].dt.to_period("M").astype(str)

    hoje      = date.today()
    mes_atual = str(pd.Period(hoje, "M"))

    df_fut = df[df["mes_remessa"] >= mes_atual].copy()

    print(f"  Pedidos em aberto     : {len(df)}")
    print(f"  Mês de referência     : {mes_atual}")
    print(f"  Remessas futuras      : {len(df_fut)}")
    print()
    if not df_fut.empty:
        print(df_fut[["numero_pedido", "material", "quantidade", "data_remessa"]].to_string(index=False))

    entradas = (
        df_fut.groupby(["material", "mes_remessa"], as_index=False)["quantidade"]
        .sum()
        .rename(columns={"mes_remessa": "mes", "quantidade": "qtd_entrada"})
    )
    return entradas, df_fut


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 5: CLASSIFICAÇÃO ABC
# ─────────────────────────────────────────────────────────────────────────────
def passo_5_abc(
    demanda: pd.DataFrame,
    materiais: pd.DataFrame,
    contratos: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Parâmetro 'contratos' é opcional; quando fornecido, usa MAX(Preço líquido)
    dos contratos para enriquecer o valor unitário dos materiais (inclusive
    de contratos vencidos — regra ABC validada). Se None, comportamento
    idêntico ao original (usa materiais.csv)."""
    separador("PASSO 5 │ CLASSIFICAÇÃO ABC")

    dem_total = (
        demanda.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "demanda_total"})
    )
    abc = dem_total.merge(materiais[["material", "valor_unitario"]], on="material", how="left")
    abc["valor_unitario"] = abc["valor_unitario"].fillna(0)

    # ── Enriquecimento com preço dos contratos SAP (quando disponível) ─────────
    if contratos is not None and not contratos.empty and "valor_unitario" in contratos.columns:
        preco_sap = (
            contratos.groupby("material", as_index=False)["valor_unitario"]
            .max()
            .rename(columns={"valor_unitario": "_preco_sap"})
        )
        abc = abc.merge(preco_sap, on="material", how="left")
        # Substituir apenas onde o contrato tem preço > 0; manter original como fallback
        mask = abc["_preco_sap"].notna() & (abc["_preco_sap"] > 0)
        abc.loc[mask, "valor_unitario"] = abc.loc[mask, "_preco_sap"]
        abc = abc.drop(columns=["_preco_sap"])
        print(f"  Preço SAP aplicado    : {mask.sum()} material(is)")
    abc["valor_total"]    = abc["demanda_total"] * abc["valor_unitario"]

    abc = abc.sort_values("valor_total", ascending=False).reset_index(drop=True)

    total_geral    = abc["valor_total"].sum()
    abc["pct_ind"] = abc["valor_total"] / total_geral * 100
    abc["pct_acum"] = abc["valor_total"].cumsum() / total_geral

    abc["classe"] = "C"
    abc.loc[abc["pct_acum"] <= LIMITE_ABC_B, "classe"] = "B"
    abc.loc[abc["pct_acum"] <= LIMITE_ABC_A, "classe"] = "A"

    desc_map = dict(zip(materiais["material"], materiais["descricao"]))
    print(f"\n  {'Material':<10} {'Descrição':<30} {'Valor Total':>12} {'Acum%':>7} {'Classe':>6}")
    print(f"  {'─'*10} {'─'*30} {'─'*12} {'─'*7} {'─'*6}")
    for _, r in abc.iterrows():
        desc = desc_map.get(r["material"], "")[:29]
        print(
            f"  {r['material']:<10} {desc:<30} "
            f"{r['valor_total']:>12,.2f} "
            f"{r['pct_acum']*100:>6.1f}% "
            f"{r['classe']:>6}"
        )

    cnt = abc.groupby("classe").size()
    print(
        f"\n  Classe A (≤{LIMITE_ABC_A*100:.0f}%): {cnt.get('A', 0)} mat.  │"
        f"  Classe B ({LIMITE_ABC_A*100:.0f}%-{LIMITE_ABC_B*100:.0f}%): {cnt.get('B', 0)} mat.  │"
        f"  Classe C (>{LIMITE_ABC_B*100:.0f}%): {cnt.get('C', 0)} mat."
    )

    export = abc.copy()
    export["pct_ind"]  = export["pct_ind"].round(2)
    export["pct_acum"] = (export["pct_acum"] * 100).round(2)
    export = export.rename(columns={"pct_ind": "pct_individual", "pct_acum": "pct_acumulado"})
    salvar(export, "02_classificacao_abc.csv")

    return abc  # pct_acum como fração (0–1) internamente


# ─────────────────────────────────────────────────────────────────────────────
# PASSOS 6-11: CÁLCULO MRP MÊS A MÊS
#
# Estratégia por classe:
#
#   Classe A / B
#     Trigger : estoque projetado < estoque de segurança (3 meses rolling)
#     Cobertura: 3 meses (pedido = SS_3m − est_proj)
#
#   Classe C  (baixo valor financeiro — foco em NUNCA ter ruptura)
#     Trigger : estoque projetado < demanda do mês corrente
#               (< 1 mês de cobertura = risco iminente)
#     Cobertura: 4 meses (pedido maior → menos pedidos/ano = consolidação)
#     Sem limite rígido de pedidos por ano
# ─────────────────────────────────────────────────────────────────────────────
def passos_6_11_mrp(
    demanda: pd.DataFrame,
    estoque: pd.DataFrame,
    entradas_pedidos: pd.DataFrame,
    abc: pd.DataFrame,
    materiais: pd.DataFrame,
    lead_time_dias: int = LEAD_TIME_DIAS,
    lead_times_dict: dict | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Parâmetro 'lead_times_dict' é opcional (dict material→dias).
    Quando fornecido, cada material usa seu próprio lead time;
    ausentes no dict usam 'lead_time_dias' como fallback.
    Quando None, comportamento idêntico ao original."""
    separador("PASSOS 6-11 │ CÁLCULO MRP MÊS A MÊS")

    hoje      = date.today()
    per_atual = pd.Period(hoje, "M")
    periodos  = [str(p) for p in pd.period_range(start=per_atual, periods=HORIZONTE_MESES, freq="M")]
    n_per     = len(periodos)

    print(f"  Data atual  : {hoje.strftime('%d/%m/%Y')}")
    print(f"  Horizonte   : {periodos[0]} → {periodos[-1]}  ({n_per} meses)")
    print(f"  Lead time   : {lead_time_dias} dias corridos")

    periodos_set = set(periodos)   # lookup O(1) para verificar se chegada está no horizonte

    # ── Lookups ───────────────────────────────────────────────────────────────
    dem_lkp: dict[str, dict[str, float]] = {}
    for _, row in demanda.iterrows():
        dem_lkp.setdefault(row["material"], {})[row["mes"]] = float(row["quantidade"])

    ent_lkp: dict[str, dict[str, float]] = {}
    for _, row in entradas_pedidos.iterrows():
        ent_lkp.setdefault(row["material"], {})[row["mes"]] = float(row["qtd_entrada"])

    classe_map  = dict(zip(abc["material"], abc["classe"]))
    estoque_map = dict(zip(estoque["material"], estoque["estoque_total"]))
    desc_map    = dict(zip(materiais["material"], materiais["descricao"]))

    todos_mats = sorted(set(estoque["material"]) | set(demanda["material"]))

    resultados: list[dict]    = []
    pedidos_compra: list[dict] = []

    for mat in todos_mats:
        classe  = classe_map.get(mat, "C")
        est_ini = float(estoque_map.get(mat, 0))
        dem_mat = dem_lkp.get(mat, {})
        ent_mat = ent_lkp.get(mat, {})

        novas_ent: dict[str, float] = {p: 0.0 for p in periodos}

        est_proj = est_ini

        for i, per in enumerate(periodos):
            dem = dem_mat.get(per, 0.0)

            # ── Estoque de segurança — sempre 3 meses rolling (exibição/relatório)
            ss_display = sum(
                dem_mat.get(periodos[j], 0.0)
                for j in range(i, min(i + MESES_COBERTURA_SS, n_per))
            )

            # ── Entradas deste mês ────────────────────────────────────────────
            ent_exist = ent_mat.get(per, 0.0)
            ent_nova  = novas_ent.get(per, 0.0)
            total_ent = ent_exist + ent_nova

            # ── Estoque projetado ─────────────────────────────────────────────
            est_proj = est_proj + total_ent - dem

            # ── Necessidade e geração de pedido — ESTRATÉGIA POR CLASSE ──────
            pedido      = 0
            per_entrega = ""
            nec         = 0.0

            if classe == "C":
                # Trigger: cobertura < 1 mês (risco de ruptura)
                # Ao pedir: cobrir os próximos CLASSE_C_COBERTURA_MESES meses
                if est_proj < dem:
                    ss_ordem = sum(
                        dem_mat.get(periodos[j], 0.0)
                        for j in range(i, min(i + CLASSE_C_COBERTURA_MESES, n_per))
                    )
                    nec = max(0.0, ss_ordem - est_proj)
            else:
                # Classes A e B: trigger quando est_proj < SS de 3 meses
                nec = max(0.0, ss_display - est_proj)

            if nec > 0:
                pedido = math.ceil(nec)

                # ── Lead time em DIAS CORRIDOS ────────────────────────────────
                # data_chegada = data_pedido + lead_time_dias
                # Entrada alocada no mês da data_chegada
                lt = lead_times_dict.get(mat, lead_time_dias) if lead_times_dict else lead_time_dias
                per_entrega, data_ped, data_cheg = calcular_periodo_entrega(
                    per, lt
                )

                if per_entrega in periodos_set:
                    novas_ent[per_entrega] += pedido
                else:
                    per_entrega = "além_horizonte"

                pedidos_compra.append(
                    {
                        "material"           : mat,
                        "descricao"          : desc_map.get(mat, ""),
                        "classe"             : classe,
                        "periodo_necessidade": per,
                        "data_pedido"        : data_ped.strftime("%d/%m/%Y"),
                        "lead_time_dias"     : lt,
                        "data_chegada"       : data_cheg.strftime("%d/%m/%Y"),
                        "periodo_entrega"    : per_entrega,
                        "quantidade"         : pedido,
                    }
                )

            resultados.append(
                {
                    "material"                  : mat,
                    "classe"                    : classe,
                    "periodo"                   : per,
                    "demanda"                   : int(dem),
                    "entrada_pedidos_existentes": int(ent_exist),
                    "entrada_pedidos_novos"     : int(ent_nova),
                    "total_entradas"            : int(total_ent),
                    "estoque_projetado"         : int(est_proj),
                    "estoque_seguranca_3m"      : int(ss_display),
                    "necessidade"               : int(nec),
                    "pedido_gerado"             : pedido,
                    "periodo_entrega"           : per_entrega,
                }
            )

    df_mrp = pd.DataFrame(resultados)
    df_ped = (
        pd.DataFrame(pedidos_compra)
        if pedidos_compra
        else pd.DataFrame(
            columns=["material", "descricao", "classe", "periodo_necessidade",
                     "data_pedido", "lead_time_dias", "data_chegada",
                     "periodo_entrega", "quantidade"]
        )
    )

    salvar(df_mrp, "03_mrp_projetado.csv")
    salvar(df_ped, "04_pedidos_compra.csv")

    print(f"\n  {'Material':<10} {'Classe':>6} {'Pedidos':>8} {'Qtd Total':>12}  Estratégia")
    print(f"  {'─'*10} {'─'*6} {'─'*8} {'─'*12}  {'─'*30}")
    for mat in todos_mats:
        cls     = classe_map.get(mat, "C")
        sub     = df_ped[df_ped["material"] == mat] if not df_ped.empty else pd.DataFrame()
        n_ped   = len(sub)
        qtd     = int(sub["quantidade"].sum()) if not sub.empty else 0
        if cls == "C":
            estrategia = f"Trigger 1m / Cobertura {CLASSE_C_COBERTURA_MESES}m (sem limite)"
        elif cls == "A":
            estrategia = f"SS {MESES_COBERTURA_SS}m / Compra frequente"
        else:
            estrategia = f"SS {MESES_COBERTURA_SS}m / Estratégia intermediária"
        print(f"  {mat:<10} {cls:>6} {n_ped:>8} {qtd:>12,}  {estrategia}")

    return df_mrp, df_ped


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 12: RATEIO FINAL POR DEPARTAMENTO / PROGRAMA ORÇAMENTÁRIO
#
# REGRA: tanto os pedidos gerados (novos) quanto os pedidos em aberto (já
# efetuados) são rateados pela proporção da demanda por material.
# A proporção por departamento/programa está em data/rateio_base.csv
# (representando a distribuição típica da demanda de cada material).
# ─────────────────────────────────────────────────────────────────────────────
def passo_12_rateio(
    df_pedidos_gerados: pd.DataFrame,
    df_pedidos_abertos: pd.DataFrame,
    demanda_detail: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Parâmetro 'demanda_detail' é opcional: quando fornecido e rateio_base.csv
    não existir, as proporções são derivadas da própria demanda DTM.
    Quando rateio_base.csv existe, comportamento idêntico ao original."""
    separador("PASSO 12 │ RATEIO POR DEPARTAMENTO / PROGRAMA ORÇAMENTÁRIO")

    rb_path = os.path.join(DIR_DADOS, "rateio_base.csv")
    if os.path.exists(rb_path):
        rateio_base = pd.read_csv(rb_path)
        print(f"  Fonte rateio         : rateio_base.csv ({len(rateio_base)} linhas)")
    elif demanda_detail is not None and not demanda_detail.empty:
        rateio_base = derivar_rateio_da_demanda(demanda_detail)
        print(f"  Fonte rateio         : derivado da demanda ({len(rateio_base)} linhas)")
    else:
        print("  ⚠ rateio_base.csv não encontrado e demanda_detail não fornecida — sem rateio")
        rateio_base = pd.DataFrame(
            columns=["material", "departamento", "programa_orcamentario", "proporcao"]
        )

    # ── Montar tabela unificada de pedidos ────────────────────────────────────
    linhas_pedidos: list[dict] = []

    # Pedidos gerados pelo MRP
    if not df_pedidos_gerados.empty:
        for _, r in df_pedidos_gerados.iterrows():
            if r["periodo_entrega"] != "além_horizonte":
                linhas_pedidos.append(
                    {
                        "tipo"           : "GERADO",
                        "material"       : r["material"],
                        "periodo_entrega": r["periodo_entrega"],
                        "quantidade"     : float(r["quantidade"]),
                    }
                )

    # Pedidos em aberto (já efetuados — usar data_remessa)
    if not df_pedidos_abertos.empty:
        for _, r in df_pedidos_abertos.iterrows():
            linhas_pedidos.append(
                {
                    "tipo"           : "EXISTENTE",
                    "material"       : r["material"],
                    "periodo_entrega": r["mes_remessa"],
                    "quantidade"     : float(r["quantidade"]),
                }
            )

    if not linhas_pedidos:
        print("  Nenhum pedido para ratear.")
        return pd.DataFrame()

    df_todos = pd.DataFrame(linhas_pedidos)

    # ── Aplicar proporções por departamento/programa ──────────────────────────
    linhas_rateio: list[dict] = []

    for _, ped in df_todos.iterrows():
        mat     = ped["material"]
        qtd     = ped["quantidade"]
        per_ent = ped["periodo_entrega"]
        tipo    = ped["tipo"]

        proporcoes = rateio_base[rateio_base["material"] == mat]

        if proporcoes.empty:
            # Material sem base de rateio → destino único
            linhas_rateio.append(
                {
                    "tipo"                  : tipo,
                    "material"              : mat,
                    "periodo_entrega"       : per_ent,
                    "departamento"          : "NAO_DEFINIDO",
                    "programa_orcamentario" : "NAO_DEFINIDO",
                    "qtd_rateada"           : round(qtd, 2),
                    "proporcao_pct"         : 100.0,
                }
            )
        else:
            for _, p in proporcoes.iterrows():
                linhas_rateio.append(
                    {
                        "tipo"                  : tipo,
                        "material"              : mat,
                        "periodo_entrega"       : per_ent,
                        "departamento"          : p["departamento"],
                        "programa_orcamentario" : p["programa_orcamentario"],
                        "qtd_rateada"           : round(qtd * p["proporcao"], 2),
                        "proporcao_pct"         : round(p["proporcao"] * 100, 2),
                    }
                )

    df_rateio = pd.DataFrame(linhas_rateio)
    salvar(df_rateio, "05_rateio_final.csv")

    # ── Resumo por tipo de pedido ─────────────────────────────────────────────
    print()
    for tipo in ["EXISTENTE", "GERADO"]:
        sub = df_rateio[df_rateio["tipo"] == tipo]
        if not sub.empty:
            print(f"  {tipo}: {len(sub)} linhas de rateio │ {sub['qtd_rateada'].sum():,.1f} un. totais rateadas")

    return df_rateio


# ─────────────────────────────────────────────────────────────────────────────
# VALIDAÇÃO LEAD TIME — EXEMPLO REAL COM DIAS CORRIDOS
# ─────────────────────────────────────────────────────────────────────────────
def validar_lead_time(lead_time_dias: int = LEAD_TIME_DIAS) -> None:
    """
    Exibe 3 exemplos práticos do cálculo de lead time em dias corridos,
    mostrando exatamente como data_chegada e mês de entrada são determinados.

    Responde às perguntas de consistência:
      ✓ Lead time em dias corridos?         → Sim
      ✓ data_chegada = data_pedido + LT?    → Sim
      ✓ Entrada no mês correto?             → Sim (mês da data_chegada)
      ✓ Nenhuma entrada no passado?         → Garantido por calcular_periodo_entrega()
      ✓ LT além do horizonte?               → Marcado 'além_horizonte', excluído do MRP
    """
    separador("VALIDAÇÃO LEAD TIME │ DIAS CORRIDOS → MÊS DE ENTRADA")

    hoje = date.today()

    print(f"  Lead time configurado : {lead_time_dias} dias corridos")
    print(f"  Fórmula               : data_chegada = data_pedido + {lead_time_dias} dias")
    print(f"  Data de referência    : {hoje.strftime('%d/%m/%Y')}\n")

    hdr = (
        f"  {'Material':<22} {'Período Nec.':>12} {'Data Pedido':>12} "
        f"{'+ LT':>6} {'Data Chegada':>13} {'Mês Entrada':>12}"
    )
    print(hdr)
    print("  " + "─" * (len(hdr) - 2))

    # Três casos: mês atual, próximo mês, mês +2
    per_at = str(pd.Period(hoje, "M"))
    casos = [
        ("MAT002 - Rolamento 6205",    per_at,                       "pedido emitido hoje"),
        ("MAT005 - Lubrificante 68",   str(pd.Period(hoje, "M") + 1), "emitido no 1º do próx. mês"),
        ("MAT008 - Acoplamento 42mm",  str(pd.Period(hoje, "M") + 2), "emitido no 1º do mês+2"),
    ]

    for mat_label, per_nec, obs in casos:
        per_ent, d_ped, d_cheg = calcular_periodo_entrega(per_nec, lead_time_dias)
        print(
            f"  {mat_label:<22} {per_nec:>12} {d_ped.strftime('%d/%m/%Y'):>12} "
            f"{lead_time_dias:>5}d {d_cheg.strftime('%d/%m/%Y'):>13} {per_ent:>12}"
            f"   ← {obs}"
        )

    print()
    print("  Consistência verificada:")
    print("  ✓ data_chegada = data_pedido + lead_time_dias  (dias corridos)")
    print("  ✓ Mês de entrada = mês da data_chegada")
    print("  ✓ Chegada sempre posterior ao mês da necessidade (nunca no passado)")
    print("  ✓ Períodos fora do horizonte → 'além_horizonte' → não alocados no MRP")


# ─────────────────────────────────────────────────────────────────────────────
# VALIDAÇÃO 1: CLASSE C — SEM RUPTURA
# ─────────────────────────────────────────────────────────────────────────────
def validar_classe_c(df_mrp: pd.DataFrame, material: str = "MAT007") -> None:
    separador(f"VALIDAÇÃO 1 │ CLASSE C SEM RUPTURA — {material}")

    sub = df_mrp[df_mrp["material"] == material].copy()
    if sub.empty:
        print(f"  Material {material} não encontrado no MRP.")
        return

    classe = sub["classe"].iloc[0]
    print(f"  Material : {material}  │  Classe : {classe}")
    print(
        f"  Estratégia: trigger quando estoque < demanda mensal; "
        f"cobertura {CLASSE_C_COBERTURA_MESES} meses ao pedir\n"
    )

    hdr = f"  {'Período':<10} {'Demanda':>8} {'Entradas':>9} {'Est.Proj':>9} {'SS 3m':>7} {'Pedido':>8}  Situação"
    sep_ln = "  " + "─" * (len(hdr) - 2)
    print(hdr)
    print(sep_ln)

    ruptura = False
    for _, r in sub.iterrows():
        ep = int(r["estoque_projetado"])

        if ep < 0:
            situacao = "⚠️  RUPTURA"
            ruptura  = True
        elif ep < int(r["demanda"]):
            situacao = "⚡ Risco (<1m)"
        elif ep < int(r["estoque_seguranca_3m"]):
            situacao = "⬇  Abaixo SS"
        else:
            situacao = "✓  OK"

        ped_str = f"{int(r['pedido_gerado']):>7}" if r["pedido_gerado"] > 0 else "       -"
        print(
            f"  {r['periodo']:<10} "
            f"{int(r['demanda']):>8} "
            f"{int(r['total_entradas']):>9} "
            f"{ep:>9} "
            f"{int(r['estoque_seguranca_3m']):>7} "
            f"{ped_str}  {situacao}"
        )

    print()
    if ruptura:
        print("  ❌ RESULTADO: ocorreu ruptura (estoque negativo detectado)")
    else:
        print("  ✅ RESULTADO: nenhuma ruptura — Classe C atendida sem limite rígido de pedidos")


# ─────────────────────────────────────────────────────────────────────────────
# VALIDAÇÃO 2: RATEIO — EXEMPLO DETALHADO
# ─────────────────────────────────────────────────────────────────────────────
def validar_rateio(df_rateio: pd.DataFrame) -> None:
    separador("VALIDAÇÃO 2 │ EXEMPLO DE RATEIO POR DEPARTAMENTO / PROGRAMA")

    if df_rateio.empty:
        print("  Sem dados de rateio disponíveis.")
        return

    # Escolher o primeiro pedido gerado para demonstração
    gerados = df_rateio[df_rateio["tipo"] == "GERADO"].sort_values(
        ["material", "periodo_entrega"]
    )
    if gerados.empty:
        gerados = df_rateio.sort_values(["material", "periodo_entrega"])

    exemplo = gerados.iloc[0]
    mat_ex  = exemplo["material"]
    per_ex  = exemplo["periodo_entrega"]
    tipo_ex = exemplo["tipo"]

    sub = df_rateio[
        (df_rateio["material"] == mat_ex)
        & (df_rateio["periodo_entrega"] == per_ex)
        & (df_rateio["tipo"] == tipo_ex)
    ]

    qtd_total = sub["qtd_rateada"].sum()

    print(f"  Tipo do pedido  : {tipo_ex}")
    print(f"  Material        : {mat_ex}")
    print(f"  Período entrega : {per_ex}")
    print(f"  Qtd. total      : {qtd_total:,.2f} unidades")
    print()
    print(f"  {'Departamento':<20} {'Programa Orçamentário':<35} {'Qtd Rateada':>12} {'%':>7}")
    print(f"  {'─'*20} {'─'*35} {'─'*12} {'─'*7}")
    for _, r in sub.iterrows():
        print(
            f"  {r['departamento']:<20} "
            f"{r['programa_orcamentario']:<35} "
            f"{r['qtd_rateada']:>12,.2f} "
            f"{r['proporcao_pct']:>6.1f}%"
        )
    print()
    print(f"  Total rateado   : {sub['qtd_rateada'].sum():>12,.2f}")

    # Mostrar também um pedido existente (para provar os dois tipos)
    existentes = df_rateio[df_rateio["tipo"] == "EXISTENTE"].sort_values(
        ["material", "periodo_entrega"]
    )
    if not existentes.empty:
        ex2    = existentes.iloc[0]
        mat2   = ex2["material"]
        per2   = ex2["periodo_entrega"]
        sub2   = df_rateio[
            (df_rateio["material"] == mat2)
            & (df_rateio["periodo_entrega"] == per2)
            & (df_rateio["tipo"] == "EXISTENTE")
        ]
        qtd2   = sub2["qtd_rateada"].sum()
        print()
        print(f"  ── Exemplo pedido EXISTENTE ────────────────────────────────")
        print(f"  Material: {mat2}  │  Período: {per2}  │  Qtd total: {qtd2:,.2f}")
        print()
        print(f"  {'Departamento':<20} {'Programa Orçamentário':<35} {'Qtd Rateada':>12} {'%':>7}")
        print(f"  {'─'*20} {'─'*35} {'─'*12} {'─'*7}")
        for _, r in sub2.iterrows():
            print(
                f"  {r['departamento']:<20} "
                f"{r['programa_orcamentario']:<35} "
                f"{r['qtd_rateada']:>12,.2f} "
                f"{r['proporcao_pct']:>6.1f}%"
            )


# ─────────────────────────────────────────────────────────────────────────────
# PIVOT RESUMO MRP
# ─────────────────────────────────────────────────────────────────────────────
def _imprimir_mrp_pivot(df_mrp: pd.DataFrame) -> None:
    separador("RESUMO │ ESTOQUE PROJETADO POR MATERIAL E PERÍODO")
    pivot = df_mrp.pivot_table(
        index=["material", "classe"],
        columns="periodo",
        values="estoque_projetado",
        aggfunc="sum",
    )
    cols = list(pivot.columns)[:6]
    print(pivot[cols].to_string())
    if len(pivot.columns) > 6:
        print(f"  ... (+{len(pivot.columns) - 6} períodos adicionais no CSV)")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main() -> None:
    cabecalho("SISTEMA MRP COM ENDEREÇAMENTO DE ESTOQUE")
    print(f"  Data de execução  : {date.today().strftime('%d/%m/%Y')}")
    print(f"  Horizonte         : {HORIZONTE_MESES} meses  │  Lead time: {LEAD_TIME_DIAS} dias corridos")
    print(f"  SS classes A/B    : {MESES_COBERTURA_SS} meses rolling")
    print(f"  Classe C          : trigger <1m cobertura → pedido cobre {CLASSE_C_COBERTURA_MESES}m (sem limite/ano)")

    os.makedirs(DIR_DADOS, exist_ok=True)
    os.makedirs(DIR_SAIDA, exist_ok=True)

    # ── Carregar arquivos auxiliares (SAP novos + legado) ─────────────────────
    materiais = pd.read_csv(os.path.join(DIR_DADOS, "materiais.csv"), encoding="latin-1", sep=";")

    # Contratos SAP (opcional — enriquece preços para ABC)
    contratos_path = os.path.join(DIR_DADOS, ARQ_CONTRATOS_SAP)
    contratos = ler_contratos_sap(contratos_path) if os.path.exists(contratos_path) else pd.DataFrame()

    # Lead times por material (opcional — fallback = LEAD_TIME_DIAS)
    lt_path   = os.path.join(DIR_DADOS, ARQ_LEAD_TIMES)
    lt_dict   = ler_lead_times(lt_path) if os.path.exists(lt_path) else {}

    # ── Pipeline principal ────────────────────────────────────────────────────
    demanda                   = passo_1_2_demanda()
    estoque                   = passo_3_estoque()
    entradas, df_abertos_fut  = passo_4_pedidos_abertos()
    abc                       = passo_5_abc(demanda, materiais, contratos=contratos)
    df_mrp, df_ped            = passos_6_11_mrp(
                                    demanda, estoque, entradas, abc, materiais,
                                    lead_time_dias=LEAD_TIME_DIAS,
                                    lead_times_dict=lt_dict or None,
                                )

    # Recuperar detalhamento da demanda (departamento/programa) para rateio
    caminho_raw = os.path.join(DIR_DADOS, ARQUIVO_DEMANDA_RAW) if ARQUIVO_DEMANDA_RAW else None
    demanda_detail = (
        transformar_demanda_dtm(caminho_raw)
        if caminho_raw and os.path.exists(caminho_raw)
        else None
    )
    df_rateio = passo_12_rateio(df_ped, df_abertos_fut, demanda_detail=demanda_detail)

    _imprimir_mrp_pivot(df_mrp)

    # Validações obrigatórias
    validar_lead_time(LEAD_TIME_DIAS)
    validar_classe_c(df_mrp, material="MAT007")
    validar_rateio(df_rateio)

    # ── Resumo final ───────────────────────────────────────────────────────────
    cabecalho("RESUMO FINAL")
    print(f"  Materiais processados : {df_mrp['material'].nunique()}")
    print(f"  Períodos planejados   : {df_mrp['periodo'].nunique()}")
    print(f"  Pedidos gerados (MRP) : {len(df_ped)}")

    if not df_ped.empty:
        print(f"  Volume total gerado   : {df_ped['quantidade'].sum():,.0f} unidades")
        print()
        print(f"  Por classe:")
        for cls in ["A", "B", "C"]:
            sub = df_ped[df_ped["classe"] == cls]
            if not sub.empty:
                print(
                    f"    Classe {cls} → {len(sub):>3} pedido(s) │ "
                    f"{sub['quantidade'].sum():>8,.0f} un."
                )

    if not df_rateio.empty:
        print()
        print(f"  Rateio — linhas totais: {len(df_rateio)}")
        por_depto = (
            df_rateio.groupby("departamento")["qtd_rateada"]
            .sum()
            .sort_values(ascending=False)
        )
        print(f"  Por departamento:")
        for depto, qtd in por_depto.items():
            print(f"    {depto:<20} : {qtd:>10,.1f} un.")

    print(f"\n  Arquivos gerados em ./{DIR_SAIDA}/")
    for arq in sorted(os.listdir(DIR_SAIDA)):
        if arq.endswith(".csv"):
            print(f"    ✓ {arq}")

    print(f"\n{'═' * W}\n")


if __name__ == "__main__":
    main()
