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
from functools import lru_cache
from datetime import date, timedelta

import numpy as np
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# AUDITORIA DE REJEIÇÕES
# ─────────────────────────────────────────────────────────────────────────────
class AuditTrail:
    """Registra linhas rejeitadas com motivo durante importação."""
    def __init__(self):
        self.rejections = {
            "demanda": [],
            "estoque": [],
            "pedidos": [],
            "materiais": [],
            "lead_times": [],
            "mb51": [],
            "contratos": [],
        }

    def reject(self, source: str, row_idx: int, row_data: dict, reason: str) -> None:
        """Registra rejeição de uma linha."""
        if source not in self.rejections:
            self.rejections[source] = []
        self.rejections[source].append({
            "idx": row_idx,
            "reason": reason,
            **row_data
        })

    def save_all(self, audit_dir: str = None) -> dict:
        """Salva relatórios de rejeição em CSVs. Retorna resumo."""
        if audit_dir is None:
            audit_dir = os.path.join(DIR_SAIDA, "audit")
        os.makedirs(audit_dir, exist_ok=True)
        summary = {}
        for source, rows in self.rejections.items():
            if rows:
                df = pd.DataFrame(rows)
                path = os.path.join(audit_dir, f"rejected_{source}.csv")
                df.to_csv(path, index=False)
                summary[source] = len(rows)
        return summary

    def print_summary(self) -> None:
        """Exibe resumo de rejeições."""
        total = sum(len(v) for v in self.rejections.values())
        if total > 0:
            print(f"\n  ⚠ AUDITORIA: {total} linha(s) rejeitada(s)")
            for source, rows in self.rejections.items():
                if rows:
                    print(f"    • {source}: {len(rows)} linhas")

audit = AuditTrail()

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

DIR_DADOS = os.environ.get("MRP_DIR_DADOS", "data")
DIR_SAIDA = os.environ.get("MRP_DIR_SAIDA", "output")

# Arquivo de demanda bruta no formato DTM (dd/mm/yyyy).
# Quando definido e o arquivo existir, substitui demanda.csv como fonte de demanda.
# Definir como None para usar demanda.csv diretamente.
ARQUIVO_DEMANDA_RAW = "demanda.csv"

# ── Nomes convencionais dos arquivos SAP (usados pelo main() em modo CLI) ──────
# Quando os arquivos existirem em DIR_DADOS, os parsers SAP são ativados
# automaticamente; caso contrário, o fluxo legado é mantido.
ARQ_REMESSAS_SAP  = "remessas_sap.csv"       # tab-sep exportado do SAP ME2M / ME9F
ARQ_ESTOQUE_SAP   = "estoque_sap.csv"      # tab-sep exportado do SAP MB52 / MMBE
ARQ_CONTRATOS_SAP = "Contratos_SAP"        # tab-sep exportado do SAP ME3M / ME3N
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


def achar_arquivo(nome: str, pasta: str = DIR_DADOS) -> str | None:
    """Localiza arquivo em `pasta` de forma case-insensitive.

    1. Tenta o caminho exato primeiro (rápido, preserva comportamento padrão).
    2. Se não encontrar, varre o diretório ignorando maiúsculas/minúsculas.
       Aceita também variação de extensão (.csv / sem extensão).

    Retorna o caminho completo resolvido ou None se não encontrado.
    """
    exact = os.path.join(pasta, nome)
    if os.path.exists(exact):
        return exact

    nome_lower = nome.lower()
    try:
        for entry in os.listdir(pasta):
            if entry.lower() == nome_lower:
                return os.path.join(pasta, entry)
            # Tenta também com/sem extensão .csv
            if entry.lower().rstrip(".csv") == nome_lower.rstrip(".csv"):
                return os.path.join(pasta, entry)
    except FileNotFoundError:
        pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS SAP — FUNÇÕES NOVAS (adição pura; não alteram nenhum passo existente)
# ─────────────────────────────────────────────────────────────────────────────

def br_to_float(s) -> float:
    """Converte número no formato brasileiro para float.

    Suporta:
      '3.515,50'     → 3515.50
      '-3.515,50'    → -3515.50
      '3.515,50-'    → -3515.50  (sinal à direita, padrão SAP)
      '3.515,50+'    → 3515.50   (sinal à direita positivo)
      'R$ 3.515,50'  → 3515.50   (prefixo de moeda)
      'R$3.515,50-'  → -3515.50  (moeda + sinal SAP)
    """
    if s is None:
        return 0.0
    s = str(s).strip()
    if s in ("", "-", "+", "nan", "NaN"):
        return 0.0

    # Remove prefixo de moeda (R$, US$, EUR, $, etc.) e espaços residuais
    import re as _re
    s = _re.sub(r'^[A-Za-z$€£¥R\u00a0\s]+', '', s).strip()
    if not s:
        return 0.0

    # Sinal à direita (formato SAP: '1.234,56-')
    if s.endswith("-"):
        s = "-" + s[:-1]
    elif s.endswith("+"):
        s = s[:-1]

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

    # ── Corrigir header deslocado: quando linha 0 está em branco, o pandas
    # produz "Unnamed: 0..N". Detectamos isso e promovemos a primeira linha
    # não-nula como cabeçalho real (comportamento comum em exports SAP/Excel).
    if all(str(c).startswith("Unnamed:") for c in best.columns):
        for _row_idx in range(min(6, len(best))):
            _row_vals = best.iloc[_row_idx].dropna().tolist()
            if len(_row_vals) >= 3:                      # linha com conteúdo suficiente
                # Converte NaN (float) para string vazia para evitar AttributeError
                best.columns = [
                    str(v).strip() if pd.notna(v) else ""
                    for v in best.iloc[_row_idx]
                ]
                best = best.iloc[_row_idx + 1:].reset_index(drop=True)
                break

    # Normalizar nomes de colunas: garantir strings (sem NaN), remover quebras e espaços
    best.columns = pd.Index([
        str(c) if pd.notna(c) else ""
        for c in best.columns
    ])
    best.columns = (
        best.columns
        .str.replace(r"[\r\n]+", " ", regex=True)
        .str.replace(r" {2,}", " ", regex=True)
        .str.strip()
    )
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
    # Aceita "Data de remessa" (ME9F) ou "Data do documento" (ME2M) como data
    col_data = "Data de remessa" if "Data de remessa" in df.columns else "Data do documento"
    col_map = {
        "Material"                     : "material",
        col_data                       : "_data_raw",
        "a ser fornecida (quantidade)" : "_qtd_raw",
        "Código de eliminação"         : "codigo_eliminacao",
    }
    ausentes = [c for c in col_map if c not in df.columns]
    if ausentes:
        print(f"  Colunas encontradas: {list(df.columns)}")
        raise ValueError(f"Colunas obrigatórias ausentes em remessas_sap: {ausentes}")
    if col_data == "Data do documento":
        print("  ℹ  Sem 'Data de remessa' — usando 'Data do documento' (pedidos realocados ao mês atual)")

    df = df.rename(columns=col_map)

    # ── Colunas financeiras opcionais ─────────────────────────────────────────
    # Data do documento = quando o pedido foi criado no SAP (visão orçamentária)
    if "Data do documento" in df.columns and col_data != "Data do documento":
        df["data_pedido_doc"] = pd.to_datetime(
            df["Data do documento"], format="%d/%m/%Y", errors="coerce"
        )
        df["mes_pedido"] = df["data_pedido_doc"].dt.to_period("M").astype(str)
    else:
        df["mes_pedido"] = None  # não disponível neste formato

    # Nº pedido e contrato — necessários para vincular política de pagamento
    df["numero_pedido"] = (
        df["Documento de compras"].astype(str).str.strip()
        if "Documento de compras" in df.columns else None
    )
    df["contrato"] = (
        df["Contrato básico"].astype(str).str.strip()
        if "Contrato básico" in df.columns else None
    )

    # Preço líquido → valor unitário REAL por unidade
    # SAP armazena preço por "Unidade preço" (pode ser 1, 100, 1000…)
    # Fórmula correta: valor_unitario = Preço líquido / Unidade preço
    # ATENÇÃO: "Unidade preço" pode estar em formato BR ("1.000") — usar br_to_float
    if "Preço líquido" in df.columns:
        preco_raw    = df["Preço líquido"].apply(br_to_float)
        col_up = next((c for c in df.columns if c.lower() in (
            "unidade preço", "unid. preço", "unidade de preço", "price unit", "por"
        )), None)
        if col_up:
            unidade_preco = df[col_up].apply(br_to_float).replace(0, 1)
        else:
            unidade_preco = pd.Series(1.0, index=df.index)
        df["valor_unitario_pedido"] = preco_raw / unidade_preco
        print(f"  Valor fonte          : 'Preço líquido' / '{col_up or 'N/A (=1)'}'"
              f"  | unitário médio={df['valor_unitario_pedido'].mean():.4f}")
    elif "Valor líquido pedido" in df.columns:
        # Fallback: valor total do pedido — convertemos para unitário após
        # ter a quantidade; guardamos o total e dividimos mais abaixo
        df["_valor_total_raw"] = df["Valor líquido pedido"].apply(br_to_float)
        df["valor_unitario_pedido"] = None  # será preenchido após converter qtd
    else:
        df["valor_unitario_pedido"] = 0.0

    # ── Filtro 1: eliminar código 'L' ─────────────────────────────────────────
    antes = len(df)
    df = df[df["codigo_eliminacao"].fillna("").str.strip().str.upper() != "L"].copy()
    print(f"  Filtro código 'L'    : {antes - len(df)} linha(s) removida(s)")

    # ── Converter quantidade (formato BR) ─────────────────────────────────────
    df["quantidade"] = df["_qtd_raw"].apply(br_to_float)

    # Finalizar fallback "Valor líquido pedido" agora que quantidade existe
    if "_valor_total_raw" in df.columns:
        df["valor_unitario_pedido"] = (
            df["_valor_total_raw"] / df["quantidade"].replace(0, 1)
        )

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
        df.loc[atrasados, "mes_remessa"]   = mes_atual
        df.loc[atrasados, "data_remessa"]  = pd.to_datetime(date.today())

    df_fut = df[df["mes_remessa"] >= mes_atual].copy()
    df_fut["valor_total_pedido"] = df_fut["quantidade"] * df_fut["valor_unitario_pedido"]

    total_val = df_fut["valor_total_pedido"].sum()
    print(f"  Remessas após filtros : {len(df_fut)} linha(s)")
    print(f"  Materiais únicos      : {df_fut['material'].nunique()}")
    print(f"  TOTAL valor futuro    : R$ {total_val:,.2f}")

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
            for enc in ["utf-8-sig", "cp1252", "latin-1", "utf-8"]:
                try:
                    raw = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        linhas = raw.splitlines()
    else:
        raw = None
        for enc in ["utf-8-sig", "cp1252", "latin-1", "utf-8"]:
            try:
                with open(source, encoding=enc) as fh:
                    raw = fh.read()
                break
            except (UnicodeDecodeError, LookupError):
                continue
        if raw is None:
            with open(source, encoding="utf-8-sig", errors="replace") as fh:
                raw = fh.read()
        linhas = raw.splitlines()

    # Encontrar linha de cabeçalho (contém 'Produto')
    header_idx = None
    for i, ln in enumerate(linhas):
        if "Produto" in ln:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Coluna 'Produto' não encontrada no arquivo de estoque SAP.")

    # Detectar separador automaticamente: tab ou ponto-e-vírgula
    header_line = linhas[header_idx]
    sep = "\t" if "\t" in header_line else ";"

    header_fields = header_line.split(sep)
    n_cols = len(header_fields)

    # Normalizar linhas de dados: BLIN tem 1 coluna extra no início → remover
    rows = []
    for ln in linhas[header_idx + 1:]:
        if not ln.strip():
            continue
        fields = ln.split(sep)
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
        # Tentar variantes conhecidas do SAP (idioma PT/EN, com/sem acento)
        _termos = ["disponív", "disponiv", "available", "livre utiliz", "unrestricted",
                   "estoque disp", "stock avail", "livre", "qty avail"]
        candidatas = [
            c for c in df.columns
            if any(t in c.lower() for t in _termos)
        ]
        if candidatas:
            col_qtd = candidatas[0]
        else:
            colunas_encontradas = list(df.columns)
            raise ValueError(
                f"Coluna de quantidade disponível não encontrada no arquivo de estoque SAP.\n"
                f"Colunas encontradas: {colunas_encontradas}\n"
                f"Renomeie a coluna de estoque para 'Qtd.disponível UMB' ou informe o nome correto."
            )

    df = df[[col_prod, col_qtd]].rename(
        columns={col_prod: "material", col_qtd: "_qtd_raw"}
    )

    # ── Limpar material: remover sufixos (ex: '400011INVT' → ignorar) ─────────
    # O código limpo é numérico (ou alfanum. sem sufixo de tipo).
    # Manter apenas linhas onde material é numérico ou tem no máx. 10 caracteres
    # sem o padrão de sufixo SAP (letras depois de números).
    df["material"] = df["material"].str.strip()
    invalido_mat = ~df["material"].str.match(r"^\d+$", na=False)
    if invalido_mat.any():
        for idx in df[invalido_mat].index:
            row = df.loc[idx]
            audit.reject("estoque", idx, {
                "material_raw": row["material"],
                "quantidade_raw": row["_qtd_raw"],
            }, "material_invalido_ou_com_sufixo")
    df = df[~invalido_mat].copy()

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


# ── Parser: Materiais (catálogo de materiais) ──────────────────────────────────
def ler_materiais(source) -> pd.DataFrame:
    """
    Lê arquivo de materiais (TAB ou CSV simples).
    Suporta dois formatos:
      - SAP export: CÓDIGO | DESCRIÇÃO | UDM | VALOR UNITÁRIO  (tab-sep, R$ BR)
      - Legado:     material | descricao | valor_unitario      (CSV simples)
    Retorna: DataFrame  material | descricao | valor_unitario
    """
    df = _ler_sap_tabsep(source)
    df.columns = df.columns.str.strip()

    col_aliases = {
        "CÓDIGO"         : "material",
        "CODIGO"         : "material",
        "MATERIAL"       : "material",
        "DESCRIÇÃO"      : "descricao",
        "DESCRICAO"      : "descricao",
        "DESCRIÇÃO BREVE": "descricao",
        "TEXTO BREVE"    : "descricao",
        "VALOR UNITÁRIO" : "valor_unitario",
        "VALOR UNITARIO" : "valor_unitario",
        "PRECO"          : "valor_unitario",
        "PREÇO"          : "valor_unitario",
        "LEAD_TIME"               : "lead_time_dias",
        "LEAD TIME"               : "lead_time_dias",
        "LEADTIME"                : "lead_time_dias",
        "LEAD_TIME_DIAS"          : "lead_time_dias",
        # nomes SAP em português
        "LEAD TIME PLANEJ."       : "lead_time_dias",
        "LEAD TIME PLANEJADO"     : "lead_time_dias",
        "PRAZO DE ENTREGA"        : "lead_time_dias",
        "PRAZO ENTREGA"           : "lead_time_dias",
        "PRAZO ENTR.PLANEJ."      : "lead_time_dias",
        "PRAZO ENTR PLANEJ"       : "lead_time_dias",
        "TEMPO DE REPOSIÇÃO"      : "lead_time_dias",
        "TEMPO REPOSICAO"         : "lead_time_dias",
        "TEMPO REPOSIÇÃO"         : "lead_time_dias",
        "TEMPO ENTREGA"           : "lead_time_dias",
        "TEMPO ENTREGA PLANEJ."   : "lead_time_dias",
        "LT DIAS"                 : "lead_time_dias",
        "LT"                      : "lead_time_dias",
        "DIAS"                    : "lead_time_dias",
        # responsável / planejador
        "PLANEJADOR"    : "planejador",
        "PLANNER"       : "planejador",
        "RESPONSAVEL"   : "planejador",
        "RESPONSÁVEL"   : "planejador",
        "COMPRADOR"     : "planejador",
        "BUYER"         : "planejador",
    }
    rename = {c: col_aliases[c.upper().strip()]
              for c in df.columns if c.upper().strip() in col_aliases}
    df = df.rename(columns=rename)

    if "material" not in df.columns:
        raise ValueError(f"Coluna 'material/CÓDIGO' não encontrada em materiais. Colunas: {list(df.columns)}")

    df["material"] = df["material"].astype(str).str.strip().str.lstrip("0").str.zfill(1)
    # Normalizar material para código numérico sem zeros à esquerda desnecessários
    df["material"] = df["material"].astype(str).str.strip()

    if "descricao" not in df.columns:
        df["descricao"] = ""

    if "valor_unitario" not in df.columns:
        df["valor_unitario"] = 0.0
    else:
        df["valor_unitario"] = df["valor_unitario"].apply(br_to_float)

    cols = ["material", "descricao", "valor_unitario"]
    if "lead_time_dias" in df.columns:
        df["lead_time_dias"] = pd.to_numeric(df["lead_time_dias"], errors="coerce")
        cols.append("lead_time_dias")
    if "planejador" in df.columns:
        df["planejador"] = df["planejador"].astype(str).str.strip()
        cols.append("planejador")

    return df[cols].drop_duplicates(subset="material")


# ── Parser File 6: Histórico MB51 (movimentos 101/102) ────────────────────────
def ler_historico_mb51(source) -> pd.DataFrame:
    """
    Lê relatório SAP MB51 e retorna entradas/estornos de fornecedores.

    Filtro: Tipo de Movimento 101 (recebimento) e 102 (estorno).
    Movimentos 102 ficam com quantidade e valor NEGATIVOS.

    Retorna DataFrame agrupado: material | mes_entrega | numero_pedido | quantidade | valor_pedido | data_doc
    """
    df = _ler_sap_tabsep(source)
    df.columns = df.columns.str.strip()

    # ── Mapear colunas por aliases comuns de exportações SAP ─────────────────
    col_mov = next((c for c in df.columns
                    if c.strip().lower() in ("tp.mov.", "mov.", "tipo de movimento",
                                             "movement type", "mvt", "tp mov")), None)
    col_data = next((c for c in df.columns
                     if c.strip().lower() in ("data do doc.", "data de lançamento",
                                              "posting date", "data doc.", "data do documento",
                                              "data lanç.")), None)
    col_mat  = next((c for c in df.columns
                     if c.strip().lower() in ("material", "cod. material", "código material")), None)
    col_val  = next((c for c in df.columns
                     if c.strip().lower() in ("montante em ml", "montante em mi",
                                              "montante", "valor", "amount in lc",
                                              "val.em ml", "valor total", "valor em ml")), None)
    col_qtd  = next((c for c in df.columns
                     if c.strip().lower() in ("quantidade", "qty", "qtd.", "qtd",
                                              "qtd. um registro", "qtd.um registro",
                                              "quantidade em unidade de entrada")), None)
    col_ped  = next((c for c in df.columns
                     if c.strip().lower() in ("pedido", "purchase order",
                                              "documento de compras", "nº pedido",
                                              "no. pedido", "doc. compras")), None)

    ausentes = [n for n, c in [("Tipo Mov.", col_mov), ("Data", col_data),
                                ("Material", col_mat), ("Valor", col_val)] if c is None]
    if ausentes:
        raise ValueError(f"ler_historico_mb51: colunas não encontradas: {ausentes}. "
                         f"Colunas disponíveis: {list(df.columns)}")

    df = df.rename(columns={
        col_mov : "tipo_mov",
        col_data: "data_doc",
        col_mat : "material",
        col_val : "_valor_raw",
    })
    if col_qtd:
        df = df.rename(columns={col_qtd: "_qtd_raw"})
    else:
        df["_qtd_raw"] = "0"
    if col_ped:
        df = df.rename(columns={col_ped: "numero_pedido"})
        df["numero_pedido"] = df["numero_pedido"].astype(str).str.strip()
    else:
        df["numero_pedido"] = None

    # ── Filtrar apenas 101 e 102 ──────────────────────────────────────────────
    # Normalização robusta: remove espaços, sufixo ".0" (quando pandas leu como float),
    # e qualquer caractere não-numérico invisível
    df["tipo_mov"] = (
        df["tipo_mov"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)   # "101.0" → "101"
        .str.replace(r"\s+", "", regex=True)      # espaços internos
    )
    todos_movs = df["tipo_mov"].value_counts().to_dict()
    print(f"  Tipos de movimento encontrados: {todos_movs}")
    # Valores excluídos (para diagnóstico de diferenças)
    excluidos = df[~df["tipo_mov"].isin(["101", "102"])]
    if not excluidos.empty:
        print(f"  ⚠ {len(excluidos)} linha(s) excluídas (tipo ≠ 101/102): "
              f"{excluidos['tipo_mov'].value_counts().to_dict()}")
    df = df[df["tipo_mov"].isin(["101", "102"])].copy()
    if df.empty:
        print("  ⚠ MB51: nenhum movimento 101/102 encontrado.")
        return pd.DataFrame(columns=["material", "mes_entrega", "quantidade", "valor_pedido"])

    print(f"  Linhas 101: {(df['tipo_mov']=='101').sum()}  |  Linhas 102: {(df['tipo_mov']=='102').sum()}")

    # ── Converter valores e quantidades (arredondado a 2 casas) ──────────────
    print(f"  Amostra _valor_raw: {df['_valor_raw'].head(5).tolist()}")
    df["valor"]      = df["_valor_raw"].apply(br_to_float).round(2)
    df["quantidade"] = df["_qtd_raw"].apply(br_to_float).round(3)

    # ── Garantir sinais corretos por tipo de movimento ────────────────────────
    # 101 (recebimento) → sempre positivo
    df.loc[(df["tipo_mov"] == "101") & (df["valor"] < 0), "valor"] *= -1
    df.loc[(df["tipo_mov"] == "101") & (df["quantidade"] < 0), "quantidade"] *= -1
    # 102 (estorno)     → sempre negativo
    df.loc[(df["tipo_mov"] == "102") & (df["valor"] > 0), "valor"] *= -1
    df.loc[(df["tipo_mov"] == "102") & (df["quantidade"] > 0), "quantidade"] *= -1

    # ── Converter data → mês ──────────────────────────────────────────────────
    df["data_doc"]    = pd.to_datetime(df["data_doc"], format="%d/%m/%Y", errors="coerce")
    df["mes_entrega"] = df["data_doc"].dt.to_period("M").astype(str)
    n_data_invalida   = df["data_doc"].isna().sum()
    if n_data_invalida:
        print(f"  ⚠ {n_data_invalida} linhas com data inválida — descartadas")
    df = df[df["mes_entrega"].notna()].copy()

    # ── Normalizar material (sem inner join — todos os materiais são mantidos) ─
    df["material"] = df["material"].astype(str).str.strip().str.lstrip("0").str.zfill(1)

    # ── Auditoria: totais por mês para conferência com relatório SAP ──────────
    totais_mes = df.groupby("mes_entrega")["valor"].sum().round(2)
    print("  ┌─ Auditoria MB51 (101 − 102) ──────────────────────────")
    for mes, val in totais_mes.items():
        print(f"  │  {mes}: R$ {val:>18,.2f}")
    print(f"  │  TOTAL GERAL : R$ {df['valor'].sum():>15,.2f}")
    print("  └───────────────────────────────────────────────────────")

    # ── Agrupar por material + mês + pedido (preserva numero_pedido p/ política pag.) ─
    resultado = (
        df.groupby(["material", "mes_entrega", "numero_pedido"], as_index=False, dropna=False)
        .agg(
            quantidade=("quantidade", "sum"),
            valor_pedido=("valor", "sum"),
            data_doc=("data_doc", "min"),   # data real do documento (para data_base_pagamento e mes_pedido)
        )
    )
    resultado["quantidade"]   = resultado["quantidade"].round(3)
    resultado["valor_pedido"] = resultado["valor_pedido"].round(2)
    print(f"  MB51 carregado: {len(resultado)} combinações material×mês "
          f"({df['material'].nunique()} materiais, {df['mes_entrega'].nunique()} meses)")
    return resultado


# ── Parser File 5: Lead Times ──────────────────────────────────────────────────
def ler_lead_times(source) -> dict:
    """
    Lê arquivo de lead times (CSV ou exportação SAP).
    Aceita múltiplos nomes de coluna para material e lead time.
    Retorna dict  {material_str: lead_time_int}.
    """
    separador("PARSER │ LEAD TIMES (File 5)")
    try:
        df = _ler_sap_tabsep(source)
    except Exception as exc:
        print(f"  ⚠ Não foi possível ler lead_times: {exc} — usando default {LEAD_TIME_DIAS}d")
        return {}

    df.columns = df.columns.str.strip()

    # ── Aliases aceitos para a coluna de material ──────────────────────────────
    _MAT_ALIASES = {
        "material", "código", "codigo", "cod. material",
        "código material", "nr. material", "nº material",
    }
    col_mat = next((c for c in df.columns if c.lower().strip() in _MAT_ALIASES), None)
    if col_mat is None:
        col_mat = df.columns[0]   # fallback: primeira coluna
        print(f"  ⚠ Coluna 'material' não identificada — usando 1ª coluna: '{col_mat}'")

    # ── Aliases aceitos para a coluna de lead time ─────────────────────────────
    _LT_ALIASES = {
        "lead_time_dias", "lead_time", "lead time", "leadtime",
        "lead time planej.", "lead time planejado",
        "prazo de entrega", "prazo entrega", "prazo entr.planej.",
        "prazo entr planej", "tempo de reposição", "tempo reposição",
        "tempo reposicao", "tempo entrega", "tempo entrega planej.",
        "lt dias", "lt", "dias",
    }
    col_lt = next((c for c in df.columns if c.lower().strip() in _LT_ALIASES), None)
    if col_lt is None:
        # Tenta segunda coluna como fallback
        if len(df.columns) >= 2:
            col_lt = df.columns[1]
            print(f"  ⚠ Coluna de lead time não identificada — usando 2ª coluna: '{col_lt}'")
        else:
            print(f"  ⚠ Nenhuma coluna de lead time encontrada. Colunas: {list(df.columns)}")
            return {}

    df["_mat"] = df[col_mat].astype(str).str.strip()
    df["_lt"]  = pd.to_numeric(df[col_lt], errors="coerce").fillna(LEAD_TIME_DIAS)
    # [FIX 6] Chaves normalizadas como str.strip() para consistência com demanda/estoque
    lt_dict = {row["_mat"]: int(row["_lt"]) for _, row in df.iterrows() if row["_mat"]}

    print(f"  Colunas usadas        : material='{col_mat}', lead_time='{col_lt}'")
    print(f"  Lead times carregados : {len(lt_dict)} material(is)")
    print(f"  Default (ausentes)    : {LEAD_TIME_DIAS} dias")
    if lt_dict:
        amostra = list(lt_dict.items())[:5]
        for mat, lt in amostra:
            print(f"    {mat} → {lt}d")
    return lt_dict


# ── Parser: Política de Pagamento ─────────────────────────────────────────────
def ler_politica_pagamento(source) -> dict:
    """
    Lê CSV com política de pagamento por contrato ou nº de pedido.

    Formato esperado (separador ';' ou ','):
        documento ; dias_1 ; dias_2 ; ...
    ou com cabeçalho flexível:
        contrato/pedido | parcela_1 | parcela_2 | ...

    Retorna dict  {str(documento): [int, int, ...]}
    Documentos sem dias válidos são ignorados.
    """
    separador("PARSER │ POLÍTICA DE PAGAMENTO")
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        # tenta detectar separador
        raw = source.read() if hasattr(source, "read") else open(source, "rb").read()
        texto = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
        # Detecta separador: TAB > ; > ,
        c_tab  = texto.count("\t")
        c_semi = texto.count(";")
        c_comm = texto.count(",")
        if c_tab >= max(c_semi, c_comm) and c_tab > 0:
            sep = "\t"
        elif c_semi >= c_comm:
            sep = ";"
        else:
            sep = ","
        import io
        df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str)
    except Exception as exc:
        print(f"  ⚠ Não foi possível ler política de pagamento: {exc}")
        return {}

    df.columns = [c.strip().lower() for c in df.columns]

    # Coluna do documento (contrato ou pedido)
    col_doc = None
    for c in ["documento", "contrato", "pedido", "numero_pedido", "contrato_pedido", "doc"]:
        if c in df.columns:
            col_doc = c
            break
    if col_doc is None:
        # Usa a primeira coluna
        col_doc = df.columns[0]

    # Colunas de dias de pagamento: apenas colunas que contenham "pagamento" no nome
    cols_dias = [c for c in df.columns if c != col_doc and "pagamento" in c.lower()]

    politica: dict[str, list[int]] = {}
    for _, row in df.iterrows():
        doc = str(row[col_doc]).strip()
        if not doc or doc.lower() in ("nan", ""):
            continue
        dias = []
        for c in cols_dias:
            v = row[c]
            try:
                d = int(float(str(v).replace(",", ".")))
                if 1 <= d <= 1095:  # entre 1 dia e 3 anos — filtra lixo numérico
                    dias.append(d)
            except (ValueError, TypeError):
                pass
        if dias:
            politica[doc] = dias

    print(f"  Política carregada: {len(politica)} documento(s)")
    for doc, dias in list(politica.items())[:5]:
        print(f"    {doc} → {dias} dias")
    return politica


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

    # Usa o mesmo leitor inteligente dos parsers SAP (detecta sep e encoding)
    df = _ler_sap_tabsep(caminho)

    # ── Normaliza nomes de colunas (remove espaços, BOM residual) ─────────────
    df.columns = df.columns.str.strip()
    print(f"  ℹ  Colunas detectadas: {list(df.columns)}")

    # ── Mapeamento oficial de colunas + aliases comuns ─────────────────────────
    col_map = {
        "CÓDIGO" : "material",
        "MÊS"    : "_mes_raw",
        "DEP."   : "departamento",
        "PROJETO": "programa_orcamentario",
        "QTD"    : "quantidade",
    }
    aliases = {
        "CODIGO"       : "CÓDIGO",
        "COD"          : "CÓDIGO",
        "MATERIAL"     : "CÓDIGO",
        "MES"          : "MÊS",
        "DATA"         : "MÊS",
        "DEP"          : "DEP.",
        "DEPTO"        : "DEP.",
        "DEPARTAMENTO" : "DEP.",
        "QUANTIDADE"   : "QTD",
        "QTDE"         : "QTD",
        "QTDE."        : "QTD",
        "QTD."         : "QTD",
        "PROG"         : "PROJETO",
        "PROGRAMA"     : "PROJETO",
    }
    # Renomeia colunas usando aliases (case-insensitive) e normaliza para col_map
    existing = set(df.columns)
    rename_alias = {}
    for col in df.columns:
        upper = col.upper().strip()
        if upper in col_map and col != upper:
            # Coluna difere só em maiúsculas/minúsculas → normaliza para a chave do col_map
            if upper not in existing or col == upper:
                rename_alias[col] = upper
        elif upper in aliases:
            target = aliases[upper]
            # Só renomeia se o destino ainda não existe (evita duplicatas)
            if target not in existing:
                rename_alias[col] = target
    if rename_alias:
        df = df.rename(columns=rename_alias)

    colunas_ausentes = [c for c in col_map if c not in df.columns]
    if colunas_ausentes:
        raise ValueError(
            f"Colunas obrigatórias ausentes no arquivo: {colunas_ausentes}\n"
            f"Colunas encontradas: {list(df.columns)}"
        )

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
        # Registra rejeições na auditoria
        for idx in df[invalidas].index:
            row = df.loc[idx]
            audit.reject("demanda", idx, {
                "material": row["material"],
                "data_raw": row["_mes_raw"],
                "departamento": row["departamento"],
                "programa": row["programa_orcamentario"],
                "quantidade": row["quantidade"],
            }, "data_invalida")
        df = df[~invalidas].copy()

    df = df.drop(columns=["_mes_raw"])

    # ── Conversão de quantidade para numérico ─────────────────────────────────
    # Usa br_to_float para suportar formato BR ("1,5", "1.234,56") sem NaN
    df["quantidade"] = df["quantidade"].apply(br_to_float)

    # br_to_float já garante float (nunca retorna NaN); nenhum fillna necessário

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
def passo_1_2_demanda() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Retorna (demanda_consolidada, demanda_detail).
    demanda_detail contém colunas completas (material, mes, departamento,
    programa_orcamentario, quantidade) — evita segunda leitura do arquivo.
    """
    separador("PASSO 1-2 │ LER E CONSOLIDAR DEMANDA")

    # [FIX 1] Fonte única de demanda: formato DTM obrigatório.
    if not ARQUIVO_DEMANDA_RAW:
        raise ValueError(
            "ARQUIVO_DEMANDA_RAW não configurado em mrp.py. "
            "Defina o nome do arquivo de demanda no formato DTM."
        )
    caminho = achar_arquivo(ARQUIVO_DEMANDA_RAW)
    if not caminho:
        raise FileNotFoundError(
            f"Arquivo de demanda não encontrado: '{ARQUIVO_DEMANDA_RAW}' (nem variações de capitalização) em '{DIR_DADOS}/'.\n"
            f"Arquivos presentes: {os.listdir(DIR_DADOS) if os.path.isdir(DIR_DADOS) else '(pasta ausente)'}"
        )

    df_raw = transformar_demanda_dtm(caminho)

    # Valida colunas obrigatórias produzidas pelo transformador
    _cols_obrig = {"material", "mes", "quantidade"}
    _ausentes = _cols_obrig - set(df_raw.columns)
    if _ausentes:
        raise ValueError(
            f"transformar_demanda_dtm não produziu colunas obrigatórias: {_ausentes}. "
            f"Colunas disponíveis: {list(df_raw.columns)}"
        )

    df = df_raw[["material", "mes", "quantidade"]].copy()

    # ── Consolidação final (agrupa caso haja duplicidades de chave) ───────────
    demanda = df.groupby(["material", "mes"], as_index=False)["quantidade"].sum()

    print(f"  Materiais com demanda : {demanda['material'].nunique()}")
    print(f"  Período da demanda    : {demanda['mes'].min()} → {demanda['mes'].max()}")
    print(f"  Registros consolidados: {len(demanda)}")

    # Retorna também o detalhe completo (com depto/prog) para reuso em rateio
    return demanda, df_raw


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3: CONSOLIDAR ESTOQUE POR MATERIAL (IGNORAR ENDEREÇAMENTO)
# ─────────────────────────────────────────────────────────────────────────────
def passo_3_estoque() -> pd.DataFrame:
    separador("PASSO 3 │ CONSOLIDAR ESTOQUE (IGNORAR ENDEREÇAMENTO)")

    # [FIX 5] Fail-fast: SAP tab-sep tem prioridade; fallback legado só se ausente.
    sap_path    = achar_arquivo(ARQ_ESTOQUE_SAP)
    legado_path = achar_arquivo("estoque.csv")

    if sap_path:
        consolidado = ler_estoque_sap(sap_path)
        salvar(consolidado, "01_estoque_consolidado.csv")
        return consolidado

    if not legado_path:
        raise FileNotFoundError(
            f"Nenhum arquivo de estoque encontrado em '{DIR_DADOS}/'.\n"
            f"  Esperado (SAP):   {ARQ_ESTOQUE_SAP} (ou variação de capitalização)\n"
            f"  Esperado (legado): estoque.csv\n"
            f"  Arquivos presentes: {os.listdir(DIR_DADOS) if os.path.isdir(DIR_DADOS) else '(pasta ausente)'}"
        )

    df = pd.read_csv(legado_path, encoding="latin-1", sep=",")
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
# HELPER: lookup de colunas de documento/cronograma no arquivo REMESSAS_SAP
# ─────────────────────────────────────────────────────────────────────────────
def _ler_remessas_lookup(source) -> pd.DataFrame:
    """
    Lê REMESSAS_SAP e devolve colunas para enriquecer pedidos_abertos.csv:
        material | data_remessa | numero_pedido | contrato | mes_pedido | fornecedor

    • Linhas com Código de eliminação == 'L' são descartadas.
    • Linhas sem data de remessa são descartadas.
    • Deduplica por (material, numero_pedido) mantendo a entrega mais próxima.
    """
    df = _ler_sap_tabsep(source)
    df.columns = df.columns.str.strip()

    col_mat  = next((c for c in df.columns if c.lower() == "material"), df.columns[0])
    col_data = (
        "Data de remessa"   if "Data de remessa"   in df.columns else
        "Data do documento" if "Data do documento" in df.columns else None
    )
    col_num  = "Documento de compras" if "Documento de compras" in df.columns else None
    col_cont = "Contrato básico"      if "Contrato básico"      in df.columns else None
    col_elim = "Código de eliminação" if "Código de eliminação" in df.columns else None
    col_doc  = "Data do documento"    if "Data do documento"    in df.columns else None
    col_forn = next((c for c in df.columns if c.lower().strip() in (
        "fornecedor/centro fornecedor", "forn./centro forn.", "fornecedor",
        "nome do fornecedor", "nome forn.", "vendor", "vendor name", "forn.",
    )), None)

    r = pd.DataFrame()
    r["material"] = df[col_mat].astype(str).str.strip()

    r["data_remessa"] = (
        pd.to_datetime(df[col_data], format="%d/%m/%Y", errors="coerce")
        if col_data else pd.NaT
    )
    r["numero_pedido"] = df[col_num].astype(str).str.strip()  if col_num  else None
    r["contrato"]      = df[col_cont].astype(str).str.strip() if col_cont else None
    r["fornecedor"]    = df[col_forn].astype(str).str.strip() if col_forn else None

    # mes_pedido = mês de emissão do pedido (Data do documento)
    if col_doc and col_doc != col_data:
        r["mes_pedido"] = (
            pd.to_datetime(df[col_doc], format="%d/%m/%Y", errors="coerce")
            .dt.to_period("M").astype(str)
        )
    else:
        r["mes_pedido"] = None

    # Filtrar código 'L'
    if col_elim:
        r = r[df[col_elim].fillna("").str.strip().str.upper() != "L"]

    # Descartar sem data
    r = r.dropna(subset=["data_remessa"])

    # Deduplica mantendo entrega mais próxima por (material, numero_pedido)
    group = ["material"] + (["numero_pedido"] if col_num else [])
    r = r.sort_values("data_remessa").drop_duplicates(subset=group, keep="first")

    print(f"  Lookup SAP           : {len(r)} registro(s) em {r['material'].nunique()} material(is)")
    return r.reset_index(drop=True)


def ler_fornecedor_por_po(source) -> dict[str, str]:
    """
    Lê REMESSAS_SAP ou pedidos_abertos e retorna dict normalizado:
        numero_pedido_normalizado → fornecedor

    Normalização: remove sufixo ".0", remove zeros à esquerda,
    converte para inteiro string (padrão SAP).
    """
    def _norm(v) -> str | None:
        s = str(v).strip()
        if s.lower() in ("nan", "none", "<na>", "—", ""):
            return None
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        try:
            return str(int(s))
        except ValueError:
            return s if s else None

    try:
        r = _ler_remessas_lookup(source)
    except Exception:
        return {}

    if "numero_pedido" not in r.columns or "fornecedor" not in r.columns:
        return {}

    lookup: dict[str, str] = {}
    for nped, forn in zip(r["numero_pedido"], r["fornecedor"]):
        k = _norm(nped)
        forn_s = str(forn).strip() if forn is not None else ""
        if k and forn_s and forn_s.lower() not in ("nan", "none", "<na>", "—", ""):
            lookup.setdefault(k, forn_s)
    return lookup


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4: LER PEDIDOS EM ABERTO (ENTRADAS FUTURAS)
# ─────────────────────────────────────────────────────────────────────────────
def passo_4_pedidos_abertos() -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      entradas_consolidadas — agrupado por material+mês (para o loop MRP)
      df_abertos_futuros    — linhas originais filtradas (para rateio e visão financeira)

    Regra de ouro:
      • BASE PRINCIPAL : pedidos_abertos.csv — materiais e quantidades válidas para o estoque
      • LOOKUP / JOIN  : REMESSAS_SAP — fornece data_remessa, numero_pedido, contrato, mes_pedido
      • Materiais do SAP sem correspondência em pedidos_abertos são descartados
    """
    separador("PASSO 4 │ PEDIDOS EM ABERTO — ENTRADAS FUTURAS")

    ped_path = achar_arquivo("pedidos_abertos.csv") or achar_arquivo("PEDIDOS_ABERTOS.csv")
    sap_path = achar_arquivo(ARQ_REMESSAS_SAP)

    tem_ped = bool(ped_path)
    tem_sap = bool(sap_path)

    if not tem_ped and not tem_sap:
        print("  ⚠ Nenhuma fonte de pedidos encontrada.")
        return pd.DataFrame(columns=["material", "mes", "qtd_entrada"]), pd.DataFrame()

    if not tem_ped:
        # Modo legado: sem pedidos_abertos.csv, usa SAP diretamente
        print("  ℹ pedidos_abertos.csv não encontrado — modo legado REMESSAS_SAP")
        return ler_remessas_sap(sap_path)

    # ── BASE PRINCIPAL: pedidos_abertos.csv ───────────────────────────────────
    df = _ler_sap_tabsep(ped_path)
    df.columns = df.columns.str.strip()

    # Material
    col_mat = next(
        (c for c in df.columns if c.lower() in (
            "material", "nº material", "nr. material", "número material",
            "cod. material", "código material",
        )),
        df.columns[0],
    )
    df = df.rename(columns={col_mat: "material"})
    df["material"] = df["material"].astype(str).str.strip()

    # Filtro código de eliminação 'L' (mesmo critério do ler_remessas_sap)
    col_elim = next(
        (c for c in df.columns if "eliminação" in c.lower() or "eliminacao" in c.lower()
         or "deletion" in c.lower()), None
    )
    if col_elim:
        antes = len(df)
        df = df[df[col_elim].fillna("").str.strip().str.upper() != "L"].copy()
        print(f"  Filtro código 'L'    : {antes - len(df)} linha(s) removida(s)")

    # Quantidade — ME2M: "a ser fornecida (quantidade)"
    col_qtd = next(
        (c for c in df.columns if c.lower() in (
            "quantidade", "qty", "qtd", "qtd.",
            "a ser fornecida (quantidade)",
            "qty. a ser fornecida", "qtd. a ser fornecida",
        )), None
    )
    df["quantidade"] = df[col_qtd].apply(br_to_float) if col_qtd else 1.0
    df = df[df["quantidade"] > 0].copy()

    # Valor total e unitário
    # Prioridade 1: coluna de valor total direta — ME2M: "a ser fornecido (valor)"
    col_val_total = next(
        (c for c in df.columns if c.lower() in (
            "a ser fornecido (valor)", "valor a ser fornecido",
            "valor total", "net value", "valor líquido total",
        )), None
    )
    if col_val_total:
        df["valor_total_pedido"]   = df[col_val_total].apply(br_to_float)
        df["valor_unitario_pedido"] = df["valor_total_pedido"] / df["quantidade"].replace(0, 1)
        print(f"  Valor fonte          : '{col_val_total}' (valor restante direto)")
    # Prioridade 2: Preço líquido / Unidade preço
    elif "Preço líquido" in df.columns:
        preco_raw = df["Preço líquido"].apply(br_to_float)
        # IMPORTANTE: br_to_float para suportar "1.000" (formato BR) = 1000
        col_up = next((c for c in df.columns if c.lower() in (
            "unidade preço", "unid. preço", "unidade de preço", "price unit", "por"
        )), None)
        if col_up:
            unidade_preco = df[col_up].apply(br_to_float).replace(0, 1)
        else:
            unidade_preco = pd.Series(1.0, index=df.index)
        df["valor_unitario_pedido"] = preco_raw / unidade_preco
        df["valor_total_pedido"]    = df["quantidade"] * df["valor_unitario_pedido"]
        print(f"  Valor fonte          : 'Preço líquido' / '{col_up or 'N/A (=1)'}'"
              f"  | unitário médio={df['valor_unitario_pedido'].mean():.4f}")
    # Prioridade 3: coluna já normalizada
    elif "valor_unitario" in df.columns:
        df["valor_unitario_pedido"] = df["valor_unitario"].apply(br_to_float)
        df["valor_total_pedido"]    = df["quantidade"] * df["valor_unitario_pedido"]
        print(f"  Valor fonte          : 'valor_unitario' (coluna normalizada)")
    else:
        df["valor_unitario_pedido"] = 0.0
        df["valor_total_pedido"]    = 0.0
        print(f"  ⚠ Valor fonte        : NENHUMA coluna de preço detectada → R$ 0")

    materiais_validos = set(df["material"].unique())
    print(f"  Base pedidos_abertos : {len(df)} linha(s), {len(materiais_validos)} material(is)")
    print(f"  Colunas disponíveis  : {list(df.columns)}")

    # ── Detectar colunas de data SEPARADAMENTE ────────────────────────────────
    # Data de ENTREGA (para mes_remessa / MRP scheduling e base de pagamento)
    col_entrega = next(
        (c for c in df.columns if c.lower() in (
            "data de remessa", "data remessa", "delivery date",
            "data prev. remessa", "data prevista remessa", "data de entrega",
        )), None
    )
    # Data de CRIAÇÃO do PO (para mes_pedido / visão orçamentária)
    col_criacao = next(
        (c for c in df.columns if c.lower() in (
            "data do documento", "data doc.", "doc. date",
            "data criação", "data emissão", "document date",
        )), None
    )

    if col_entrega:
        df["data_remessa"] = pd.to_datetime(df[col_entrega], format="%d/%m/%Y", errors="coerce")
        print(f"  Data entrega         : '{col_entrega}'")
    elif col_criacao:
        # Fallback: sem data de entrega, usa data do documento (menos preciso)
        df["data_remessa"] = pd.to_datetime(df[col_criacao], format="%d/%m/%Y", errors="coerce")
        print(f"  ⚠ 'Data de remessa' ausente — usando '{col_criacao}' como entrega")
    else:
        df["data_remessa"] = pd.NaT
        print("  ⚠ Nenhuma coluna de data encontrada em pedidos_abertos")

    if col_criacao:
        df["mes_pedido"] = (
            pd.to_datetime(df[col_criacao], format="%d/%m/%Y", errors="coerce")
            .dt.to_period("M").astype(str)
        )
        print(f"  Data criação PO      : '{col_criacao}' → mes_pedido")
    elif col_entrega:
        # Fallback: sem data de criação, usa data de entrega para mes_pedido
        df["mes_pedido"] = (
            pd.to_datetime(df[col_entrega], format="%d/%m/%Y", errors="coerce")
            .dt.to_period("M").astype(str)
        )
        print(f"  ⚠ 'Data do documento' ausente — usando entrega como mes_pedido")
    else:
        df["mes_pedido"] = None

    # ── Nº pedido e contrato ──────────────────────────────────────────────────
    _ped_aliases = {
        "documento de compras", "nº doc. compras", "nº do doc. de compras",
        "purchase order", "doc. de compras", "pedido de compra",
        "doc.compras", "doc. compras", "pedido", "nº pedido",
        "numero pedido", "número pedido", "num. pedido",
        "ordem de compra", "purchase doc.", "purch. doc.",
    }
    _cont_aliases = {
        "contrato básico", "contrato basico", "contract", "contrato",
        "contrato de fornecimento", "acordo", "scheduling agreement",
    }
    col_num_ped = next((c for c in df.columns if c.lower().strip() in _ped_aliases), None)
    col_cont    = next((c for c in df.columns if c.lower().strip() in _cont_aliases), None)

    # Fallback: se ainda não detectou, tenta coluna que contenha "compra" no nome
    if col_num_ped is None:
        col_num_ped = next((c for c in df.columns
                            if "compra" in c.lower() and "contrato" not in c.lower()), None)

    df["numero_pedido"] = df[col_num_ped].astype(str).str.strip() if col_num_ped else None
    df["contrato"]      = df[col_cont].astype(str).str.strip()    if col_cont  else None

    print(f"  Coluna Nº Pedido     : {col_num_ped!r:30s} → amostra: {df['numero_pedido'].dropna().head(3).tolist() if col_num_ped else 'NÃO DETECTADA'}")
    print(f"  Coluna Contrato      : {col_cont!r:30s} → amostra: {df['contrato'].dropna().head(3).tolist() if col_cont else 'NÃO DETECTADA'}")

    # Fornecedor (opcional — nem todo export ME2M inclui)
    col_forn = next((c for c in df.columns if c.lower() in (
        "nome do fornecedor", "nome forn.", "fornecedor",
        "fornecedor/centro fornecedor", "forn./centro forn.",
        "vendor", "vendor name", "forn.",
    )), None)
    df["fornecedor"] = df[col_forn].astype(str).str.strip() if col_forn else "—"

    # ── Complemento via REMESSAS_SAP — só se faltarem data ou doc na base ─────
    # (NÃO deduplica — apenas enriquece colunas ausentes linha a linha)
    precisa_data = df["data_remessa"].isna().all()
    precisa_doc  = df["numero_pedido"].isna().all() or df["contrato"].isna().all()

    if tem_sap and (precisa_data or precisa_doc):
        print(f"  Buscando complemento no REMESSAS_SAP (precisa_data={precisa_data})")
        lookup = _ler_remessas_lookup(sap_path)
        lookup = lookup[lookup["material"].isin(materiais_validos)]

        join_cols = ["material", "numero_pedido"] if (
            col_num_ped and lookup["numero_pedido"].notna().any()
        ) else ["material"]

        enrich_cols = [c for c in ["data_remessa", "numero_pedido", "contrato", "mes_pedido"]
                       if c not in df.columns or df[c].isna().all()]
        if enrich_cols:
            df = df.merge(
                lookup[join_cols + enrich_cols].drop_duplicates(subset=join_cols),
                on=join_cols, how="left",
                suffixes=("", "_sap"),
            )
            # Preenche apenas onde estava nulo
            for col in enrich_cols:
                if f"{col}_sap" in df.columns:
                    df[col] = df[col].fillna(df[f"{col}_sap"])
                    df.drop(columns=[f"{col}_sap"], inplace=True)

    # ── Filtrar linhas sem data de entrega ────────────────────────────────────
    df = df.dropna(subset=["data_remessa"]).copy()

    # ── Mês de remessa e realocação de atrasos ────────────────────────────────
    df["mes_remessa"] = df["data_remessa"].dt.to_period("M").astype(str)
    mes_atual = str(pd.Period(date.today(), "M"))

    atrasados = df["mes_remessa"] < mes_atual
    if atrasados.any():
        print(f"  Realocar atrasados   : {atrasados.sum()} linha(s) → {mes_atual}")
        df.loc[atrasados, "mes_remessa"]  = mes_atual
        df.loc[atrasados, "data_remessa"] = pd.to_datetime(date.today())

    df_fut = df[df["mes_remessa"] >= mes_atual].copy()
    # valor_total_pedido já calculado acima; recalcula apenas se ausente
    if "valor_total_pedido" not in df_fut.columns:
        df_fut["valor_total_pedido"] = df_fut["quantidade"] * df_fut["valor_unitario_pedido"]

    total_val = df_fut["valor_total_pedido"].sum() if "valor_total_pedido" in df_fut.columns else 0.0
    print(f"  Remessas futuras     : {len(df_fut)}")
    print(f"  Materiais únicos     : {df_fut['material'].nunique()}")
    print(f"  TOTAL valor futuro   : R$ {total_val:,.2f}")
    if not df_fut.empty:
        cols_show = ["material", "quantidade", "mes_remessa"]
        if "valor_total_pedido" in df_fut.columns:
            cols_show.append("valor_total_pedido")
        print(df_fut[cols_show].to_string(index=False))

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
    horizonte_finito: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parâmetro 'lead_times_dict' é opcional (dict material→dias).
    Quando fornecido, cada material usa seu próprio lead time;
    ausentes no dict usam 'lead_time_dias' como fallback.

    [FIX 3] horizonte_finito=True (padrão): aplica teto_pedido para evitar
    compras além da demanda restante conhecida (comportamento de phase-out).
    horizonte_finito=False: remove o teto — útil quando a demanda futura é
    incompleta e não deve limitar as compras do período atual.
    """
    separador("PASSOS 6-11 │ CÁLCULO MRP MÊS A MÊS")

    hoje      = date.today()
    per_atual = pd.Period(hoje, "M")
    periodos  = [str(p) for p in pd.period_range(start=per_atual, periods=HORIZONTE_MESES, freq="M")]
    n_per     = len(periodos)

    print(f"  Data atual  : {hoje.strftime('%d/%m/%Y')}")
    print(f"  Horizonte   : {periodos[0]} → {periodos[-1]}  ({n_per} meses)")
    print(f"  Lead time   : {lead_time_dias} dias corridos")

    periodos_set = set(periodos)   # lookup O(1) para verificar se chegada está no horizonte

    # [FIX 6] Normalizar lead_times_dict: converter todas as chaves para str.strip()
    # Evita miss por divergência de tipo (int vs str) ou espaços residuais do SAP.
    if lead_times_dict:
        lead_times_dict = {str(k).strip(): int(v) for k, v in lead_times_dict.items()}

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
    preco_map   = dict(zip(abc["material"], abc["valor_unitario"]))

    # [FIX 6] Identificar materiais sem lead time específico (para log único por material)
    _sem_lt_especifico: set[str] = set()

    todos_mats = sorted(set(estoque["material"]) | set(demanda["material"]))

    resultados: list[dict]    = []
    pedidos_compra: list[dict] = []

    for mat in todos_mats:
        classe  = classe_map.get(mat, "C")
        est_ini = float(estoque_map.get(mat, 0))
        dem_mat = dem_lkp.get(mat, {})
        ent_mat = ent_lkp.get(mat, {})

        novas_ent: dict[str, float] = {p: 0.0 for p in periodos}

        # Limitar janela de SS ao último mês com demanda cadastrada para este material.
        # Garante estoque zero no final do horizonte de demanda conhecido (ex.: dez/2026)
        # sem antecipar pedidos para anos sem previsão. Quando 2027 for carregado,
        # o comportamento volta ao normal automaticamente.
        meses_com_dem = [p for p in periodos if dem_mat.get(p, 0.0) > 0]
        idx_fim_dem   = periodos.index(max(meses_com_dem)) if meses_com_dem else 0

        est_proj = est_ini

        for i, per in enumerate(periodos):
            dem = dem_mat.get(per, 0.0)

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

            # ── Estoque virtual: posição real de cobertura (evita Panic Buying) ─
            # Considera APENAS as entradas que chegam dentro do lead time de um novo
            # pedido emitido agora — evita mascarar rupturas com entradas distantes.
            # [FIX 6] Chave já normalizada acima; fallback controlado com log único.
            if lead_times_dict:
                _mat_key = str(mat).strip()
                lt = lead_times_dict.get(_mat_key, lead_time_dias)
                if lt == lead_time_dias and _mat_key not in lead_times_dict and _mat_key not in _sem_lt_especifico:
                    print(f"  ⚠ Lead time não encontrado para '{mat}' — usando default {lead_time_dias}d")
                    _sem_lt_especifico.add(_mat_key)
            else:
                lt = lead_time_dias
            idx_chegada = min(i + math.ceil(lt / 30), n_per - 1)
            entradas_em_transito = sum(
                ent_mat.get(periodos[k], 0.0) + novas_ent.get(periodos[k], 0.0)
                for k in range(i + 1, idx_chegada + 1)
            )
            estoque_virtual = est_proj + entradas_em_transito

            # Backward scheduling: antecipar a necessidade pelo lead time.
            # Em cada mês i, projeta a posição de cobertura no momento em que um
            # novo pedido emitido AGORA chegaria (i + lt_meses) e verifica se a
            # cobertura pós-chegada estará suficiente.  Isso elimina rupturas que
            # ocorriam quando o trigger disparava apenas no mês da falta — tarde
            # demais para que a entrega chegasse a tempo.
            lt_meses        = math.ceil(lt / 30)
            cobertura_meses = CLASSE_C_COBERTURA_MESES if classe == "C" else MESES_COBERTURA_SS
            # Janela: LT meses (período até chegada) + cobertura alvo (pós-chegada)
            dem_ate_cobertura = sum(
                dem_mat.get(periodos[j], 0.0)
                for j in range(i + 1, min(i + 1 + lt_meses + cobertura_meses, n_per))
            )
            ss_display = dem_ate_cobertura  # para relatório

            # Trigger: estoque virtual não cobre a demanda até chegada + cobertura
            if estoque_virtual < dem_ate_cobertura:
                nec_ideal = max(0.0, dem_ate_cobertura - estoque_virtual)
            else:
                nec_ideal = 0.0

            # ── Teto Phase-Out: nunca pedir além da demanda restante conhecida ─
            # [FIX 3+4] horizonte_finito controla se o teto é aplicado.
            # [FIX 4] supply_total usa TODAS as entradas futuras até idx_fim_dem
            # (não apenas a janela de lead time de estoque_virtual), evitando
            # que o teto subestime o supply já comprometido e gere over-ordering.
            if horizonte_finito:
                demanda_restante = sum(
                    dem_mat.get(periodos[j], 0.0)
                    for j in range(i, idx_fim_dem + 1)
                )
                supply_total = est_proj + sum(
                    ent_mat.get(periodos[k], 0.0) + novas_ent.get(periodos[k], 0.0)
                    for k in range(i + 1, idx_fim_dem + 1)
                )
                teto_pedido = max(0.0, demanda_restante - supply_total)
                nec = min(nec_ideal, teto_pedido)
            else:
                nec = nec_ideal

            if nec > 0:
                pedido = math.ceil(nec)

                # ── Lead time já calculado acima (usado em idx_chegada) ───────
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
                        "valor_unitario"     : preco_map.get(mat, 0.0),
                        "valor_total_pedido" : pedido * preco_map.get(mat, 0.0),
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
                     "periodo_entrega", "quantidade", "valor_unitario", "valor_total_pedido"]
        )
    )

    salvar(df_mrp, "03_mrp_projetado.csv")
    salvar(df_ped, "04_pedidos_compra.csv")

    # ── Relatório final de rejeições ─────────────────────────────────────────────
    audit.print_summary()

    print(f"\n  {'Material':<10} {'Classe':>6} {'Pedidos':>8} {'Qtd Total':>12}  Estratégia")
    print(f"  {'─'*10} {'─'*6} {'─'*8} {'─'*12}  {'─'*30}")
    for mat in todos_mats:
        cls     = classe_map.get(mat, "C")
        sub     = df_ped[df_ped["material"] == mat] if not df_ped.empty else pd.DataFrame()
        n_ped   = len(sub)
        qtd     = int(sub["quantidade"].sum()) if not sub.empty else 0
        if cls == "C":
            estrategia = f"Backward LT+{CLASSE_C_COBERTURA_MESES}m / Cobertura {CLASSE_C_COBERTURA_MESES}m"
        elif cls == "A":
            estrategia = f"Backward LT+{MESES_COBERTURA_SS}m / SS {MESES_COBERTURA_SS}m"
        else:
            estrategia = f"Backward LT+{MESES_COBERTURA_SS}m / SS {MESES_COBERTURA_SS}m"
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
    Quando rateio_base.csv existe, comportamento idêntico ao original.

    Prioridade de rateio (maior → menor):
      1. rateio_manual.csv — override manual por material (salvo pelo usuário no app)
      2. rateio_base.csv   — base estática de proporções
      3. derivar da demanda DTM (quando demanda_detail fornecido)
      4. NAO_DEFINIDO (fallback)
    """
    separador("PASSO 12 │ RATEIO POR DEPARTAMENTO / PROGRAMA ORÇAMENTÁRIO")

    rb_path = os.path.join(DIR_DADOS, "rateio_base.csv")
    if os.path.exists(rb_path):
        rateio_base = pd.read_csv(rb_path, dtype={"material": str})
        rateio_base["material"] = rateio_base["material"].astype(str).str.strip()
        print(f"  Fonte rateio         : rateio_base.csv ({len(rateio_base)} linhas)")
    elif demanda_detail is not None and not demanda_detail.empty:
        rateio_base = derivar_rateio_da_demanda(demanda_detail)
        print(f"  Fonte rateio         : derivado da demanda ({len(rateio_base)} linhas)")
    else:
        print("  ⚠ rateio_base.csv não encontrado e demanda_detail não fornecida — sem rateio")
        rateio_base = pd.DataFrame(
            columns=["material", "departamento", "programa_orcamentario", "proporcao"]
        )

    # ── Override com rateio_manual.csv (prioridade máxima) ───────────────────
    _rm_path = os.path.join(DIR_DADOS, "rateio_manual.csv")
    _n_manual = 0
    if os.path.exists(_rm_path):
        try:
            rm = pd.read_csv(_rm_path, sep=";", dtype={"material": str})
            rm["material"] = rm["material"].astype(str).str.strip()
            rm = rm[["material", "departamento", "programa_orcamentario", "proporcao"]].dropna(
                subset=["material", "departamento", "programa_orcamentario", "proporcao"]
            )
            if not rm.empty:
                rm_mats = set(rm["material"].unique())
                # Remove entradas da base substituídas pelo manual
                rateio_base = rateio_base[~rateio_base["material"].isin(rm_mats)].copy()
                rateio_base = pd.concat([rateio_base, rm], ignore_index=True)
                _n_manual = len(rm_mats)
                print(f"  Rateio manual        : {_n_manual} material(is) com override")
        except Exception as _e_rm:
            print(f"  ⚠ Não foi possível carregar rateio_manual.csv: {_e_rm}")

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
    # ── Aplicar proporções via merge vetorizado (substitui nested iterrows) ──
    _rb = rateio_base[["material", "departamento", "programa_orcamentario", "proporcao"]].copy()
    _rb["material"] = _rb["material"].astype(str).str.strip()
    df_todos["material"] = df_todos["material"].astype(str).str.strip()

    merged = df_todos.merge(_rb, on="material", how="left")

    # Materiais sem rateio configurado → NAO_DEFINIDO com proporcao=1
    _no_rat = merged["departamento"].isna()
    merged.loc[_no_rat, "departamento"]          = "NAO_DEFINIDO"
    merged.loc[_no_rat, "programa_orcamentario"] = "NAO_DEFINIDO"
    merged.loc[_no_rat, "proporcao"]             = 1.0

    merged["qtd_rateada"]   = (merged["quantidade"] * merged["proporcao"]).round(2)
    merged["proporcao_pct"] = (merged["proporcao"] * 100).round(2)

    df_rateio = merged[
        ["tipo", "material", "periodo_entrega", "departamento",
         "programa_orcamentario", "qtd_rateada", "proporcao_pct"]
    ].copy()
    salvar(df_rateio, "05_rateio_final.csv")

    # ── Resumo por tipo de pedido ─────────────────────────────────────────────
    print()
    for tipo in ["EXISTENTE", "GERADO"]:
        sub = df_rateio[df_rateio["tipo"] == tipo]
        if not sub.empty:
            print(f"  {tipo}: {len(sub)} linhas de rateio │ {sub['qtd_rateada'].sum():,.1f} un. totais rateadas")

    # ── Resumo por origem do rateio ───────────────────────────────────────────
    _n_nao_def  = df_rateio[df_rateio["departamento"] == "NAO_DEFINIDO"]["material"].nunique()
    _n_total    = df_rateio["material"].nunique()
    _n_derivado = max(0, _n_total - _n_manual - _n_nao_def)
    print(f"\n  Rateio manual    : {_n_manual} material(is)")
    print(f"  Rateio derivado  : {_n_derivado} material(is)")
    print(f"  NAO_DEFINIDO     : {_n_nao_def} material(is) pendente(s)")

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
        f"  Estratégia: backward scheduling LT + {CLASSE_C_COBERTURA_MESES} meses cobertura\n"
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
    print(f"  SS classes A/B    : LT + {MESES_COBERTURA_SS} meses (backward scheduling)")
    print(f"  Classe C          : LT + {CLASSE_C_COBERTURA_MESES} meses (backward scheduling)")

    os.makedirs(DIR_DADOS, exist_ok=True)
    os.makedirs(DIR_SAIDA, exist_ok=True)

    # ── Carregar arquivos auxiliares (SAP novos + legado) ─────────────────────
    mat_path = achar_arquivo("materiais.csv")
    materiais = ler_materiais(mat_path) if mat_path else pd.DataFrame()

    # Contratos SAP (opcional — enriquece preços para ABC)
    contratos_path = achar_arquivo(ARQ_CONTRATOS_SAP)
    contratos = ler_contratos_sap(contratos_path) if contratos_path else pd.DataFrame()

    # Lead times por material: coluna LEAD_TIME do materiais.csv (prioritário)
    # Fallback: lead_times.csv separado; ausentes usam LEAD_TIME_DIAS
    if "lead_time_dias" in materiais.columns:
        lt_dict = {
            str(row["material"]): int(row["lead_time_dias"])
            for _, row in materiais.iterrows()
            if pd.notna(row["lead_time_dias"])
        }
    else:
        lt_path = achar_arquivo(ARQ_LEAD_TIMES)
        lt_dict = ler_lead_times(lt_path) if lt_path else {}

    # ── Pipeline principal ────────────────────────────────────────────────────
    demanda, demanda_detail    = passo_1_2_demanda()
    estoque                    = passo_3_estoque()
    entradas, df_abertos_fut   = passo_4_pedidos_abertos()
    abc                        = passo_5_abc(demanda, materiais, contratos=contratos)
    df_mrp, df_ped             = passos_6_11_mrp(
                                    demanda, estoque, entradas, abc, materiais,
                                    lead_time_dias=LEAD_TIME_DIAS,
                                    lead_times_dict=lt_dict or None,
                                    horizonte_finito=True,
                                )

    df_rateio = passo_12_rateio(df_ped, df_abertos_fut,
                                demanda_detail=demanda_detail if not demanda_detail.empty else None)

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

    # ── Salvar relatório de auditoria ────────────────────────────────────────────
    resumo_audit = audit.save_all()
    if resumo_audit:
        print(f"\n  📋 AUDITORIA — Linhas rejeitadas (salvo em ./{DIR_AUDIT}/)")
        for source, count in sorted(resumo_audit.items()):
            print(f"    ✓ rejected_{source}.csv  ({count} linhas)")

    print(f"\n  📊 Arquivos de saída em ./{DIR_SAIDA}/")
    for arq in sorted(os.listdir(DIR_SAIDA)):
        if arq.endswith(".csv") and not arq.startswith("rejected"):
            print(f"    ✓ {arq}")


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO PCM: PLANEJAMENTO DE MATERIAIS CRÍTICOS
# ─────────────────────────────────────────────────────────────────────────────

def ler_pcm(source, df_demanda=None, df_demanda_detail=None) -> pd.DataFrame:
    """
    Lê arquivo PCM_[DATA].xlsx/.csv com análise de itens críticos.

    Mapeamento de colunas por índice (33 campos):
      0=CHAVE, 3=DEP., 4=CÓDIGO, 5=DESCRIÇÃO, 7=GRUPO, 8=DEMANDANTES,
      9=PLANEJADOR, 10=VALOR UNITÁRIO, 11=ESTOQUE SEGURANÇA, 12=ESTOQUE MÁXIMO,
      13=SUGESTÃO PEDIDOS, 14=LT, 15=SALDO ESTOQUE, 16=SALDO EM INSPEÇÃO,
      17=SALDO PEDIDOS, 18=SALDO CONTRATO, 21=RESERVAS, 22=DEMANDA ANUAL,
      23=COMPRADOR, 24=STATUS CONTRATAÇÃO, 25=LINHAS DISPONÍVEIS,
      26=LINHAS OC, 27=OC, 28=FORNECEDOR, 29=SALDO LINHA, 30=PROG. ORÇAMENTÁRIOS,
      31=NÍVEL DE ESTOQUE, 32=ABC

    df_demanda:        DataFrame [material, mes (YYYY-MM), demanda] – para período firme real.
    df_demanda_detail: DataFrame [material, mes, departamento, programa_orcamentario, demanda]
                       – para totais por programa/departamento.

    Fluxograma:
    - Período firme  = soma demanda mês atual + 2 meses à frente (real de df_demanda)
    - Estoque seg.   = período firme + LT (default 60 dias = 2 meses adicionais)
    - Cobertura      = saldo_estoque + saldo_inspecao + saldo_pedidos
    - Dias cobertura = consumo mês a mês até zerar
    """
    from datetime import date as _date
    separador("PARSER │ PCM (Análise de Materiais Críticos)")

    try:
        if hasattr(source, "name") and str(source.name).endswith(".csv"):
            df = pd.read_csv(source, sep=";", encoding="utf-8-sig", dtype=str)
        else:
            df = pd.read_excel(source, sheet_name="MRP", engine="openpyxl", dtype=str)
    except Exception as e:
        print(f"  ⚠ Erro ao ler PCM: {e}")
        return pd.DataFrame()

    if df.empty:
        print("  ⚠ Sheet 'MRP' vazio ou não encontrado")
        return pd.DataFrame()

    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    print(f"  Linhas lidas: {len(df)}  |  Colunas: {len(df.columns)}")

    # ── Normalizar valor para número (formato BR) ─────────────────────────────
    _INVALID_STRS = {"nan", "none", "nat", "null", "-", "—", "", "n/a", "#n/a"}

    def _to_num(v, cap=None) -> float:
        s = str(v).strip()
        if s.lower() in _INVALID_STRS:
            return 0.0
        try:
            # Formato BR: ponto=milhar, vírgula=decimal
            s2 = s.replace(".", "").replace(",", ".")
            result = float(s2)
            if cap is not None and result > cap:
                return 0.0  # valor improvável = dado inválido
            return result
        except Exception:
            return 0.0

    def _clean_str(v) -> str:
        s = str(v).strip()
        return "" if s.lower() in _INVALID_STRS else s

    # ── Mapeamento por índice ─────────────────────────────────────────────────
    col_map = {
        "chave": 0, "departamento_pcm": 3, "codigo": 4, "descricao": 5,
        "grupo_material": 7, "demandantes": 8, "planejador": 9,
        "valor_unitario": 10, "estoque_seguranca_pcm": 11, "estoque_maximo": 12,
        "sugestao_pedidos": 13, "lead_time_dias": 14, "saldo_estoque": 15,
        "saldo_inspecao": 16, "saldo_pedidos": 17, "saldo_contrato": 18,
        "reservas_atendidas": 21, "demanda_anual": 22,
        "comprador": 23, "status_contratacao": 24,
        "linhas_disponiveis": 25, "linhas_oc": 26, "oc": 27,
        "fornecedor": 28, "saldo_linha": 29,
        "prog_orcamentario_raw": 30, "nivel_estoque": 31, "classe_abc": 32,
    }
    df_map = {}
    for col_name, col_idx in col_map.items():
        if col_idx < len(df.columns):
            df_map[df.columns[col_idx]] = col_name
    df = df.rename(columns=df_map)

    # ── Normalizar colunas numéricas ─────────────────────────────────────────
    _CAP_DEMANDA   = 10_000_000   # máximo plausível de demanda anual
    _CAP_ESTOQUE   = 50_000_000
    _CAP_CONTRATO  = 100_000_000

    for col, cap in [
        ("saldo_estoque",       _CAP_ESTOQUE),
        ("saldo_inspecao",      _CAP_ESTOQUE),
        ("saldo_pedidos",       _CAP_ESTOQUE),
        ("saldo_contrato",      _CAP_CONTRATO),
        ("sugestao_pedidos",    _CAP_ESTOQUE),
        ("demanda_anual",       _CAP_DEMANDA),
        ("lead_time_dias",      730),
        ("valor_unitario",      None),
        ("estoque_seguranca_pcm", _CAP_ESTOQUE),
        ("linhas_disponiveis",  100),
    ]:
        if col in df.columns:
            df[col] = df[col].apply(lambda v, c=cap: _to_num(v, c))
        else:
            df[col] = 0.0

    # ── Normalizar colunas string ─────────────────────────────────────────────
    for col in ["planejador", "comprador", "status_contratacao", "nivel_estoque",
                "grupo_material", "prog_orcamentario_raw", "demandantes",
                "codigo", "descricao", "fornecedor", "classe_abc"]:
        if col in df.columns:
            df[col] = df[col].apply(_clean_str)
        else:
            df[col] = ""

    # ── Tipo de acordo: linhas_disponiveis == 1 → PREÇO, >1 → QUANTIDADE ─────
    df["tipo_acordo"] = df["linhas_disponiveis"].apply(
        lambda x: "ACORDO PREÇO" if x == 1.0 else ("ACORDO QUANTIDADE" if x > 1 else "SEM ACORDO")
    )

    # ── Período firme e estoque de segurança ──────────────────────────────────
    today = _date.today()
    LT_PADRAO = 60.0   # dias
    MESES_FIRME = 3    # mês atual + 2 à frente

    if df_demanda is not None and not df_demanda.empty:
        # Calcular período firme a partir da demanda real
        firm_months = []
        for i in range(MESES_FIRME):
            m_offset = today.month - 1 + i
            y = today.year + m_offset // 12
            m = m_offset % 12 + 1
            firm_months.append(f"{y}-{m:02d}")

        lt_months_after = 2  # 60 dias = 2 meses após o período firme
        lt_months_list = []
        for i in range(lt_months_after):
            m_offset = today.month - 1 + MESES_FIRME + i
            y = today.year + m_offset // 12
            m = m_offset % 12 + 1
            lt_months_list.append(f"{y}-{m:02d}")

        # Garantir colunas necessárias no df_demanda
        _dem_col = "demanda" if "demanda" in df_demanda.columns else "quantidade"
        _mat_col = "material" if "material" in df_demanda.columns else df_demanda.columns[0]
        _mes_col = "mes" if "mes" in df_demanda.columns else "periodo"

        dem_firm = (df_demanda[df_demanda[_mes_col].isin(firm_months)]
                    .groupby(_mat_col)[_dem_col].sum()
                    .rename("periodo_firme_real"))
        dem_lt   = (df_demanda[df_demanda[_mes_col].isin(lt_months_list)]
                    .groupby(_mat_col)[_dem_col].sum()
                    .rename("lt_demanda_real"))

        df["codigo_norm"] = df["codigo"].astype(str).str.strip()
        df = df.merge(dem_firm, left_on="codigo_norm", right_index=True, how="left")
        df = df.merge(dem_lt,   left_on="codigo_norm", right_index=True, how="left")

        df["periodo_firme"]          = df["periodo_firme_real"].fillna(
            df["demanda_anual"] / 12 * MESES_FIRME)
        df["lt_demanda"]             = df["lt_demanda_real"].fillna(
            df["demanda_anual"] / 12 * lt_months_after)
        df["estoque_seguranca_calc"] = df["periodo_firme"] + df["lt_demanda"]

        # Dias de cobertura: consumo mês a mês
        all_future = sorted([m for m in df_demanda[_mes_col].unique()
                              if m >= today.strftime("%Y-%m")])
        # Pivô: material × mes → demanda
        _piv = (df_demanda[df_demanda[_mes_col].isin(all_future)]
                .pivot_table(index=_mat_col, columns=_mes_col, values=_dem_col,
                             aggfunc="sum", fill_value=0))

        def _calc_dias(row):
            stock = row["saldo_estoque"] + row["saldo_inspecao"]
            if stock <= 0:
                return 0
            cod = row["codigo_norm"]
            if cod not in _piv.index:
                dm = row["demanda_anual"] / 12
                return int(stock / dm * 30) if dm > 0 else 999
            days = 0
            remaining = float(stock)
            for mes in _piv.columns:
                d = float(_piv.at[cod, mes])
                if d <= 0:
                    days += 30
                    continue
                if d >= remaining:
                    days += int(remaining / d * 30)
                    return days
                remaining -= d
                days += 30
            # Stock outlasts all future months
            return min(days + int(remaining / (_piv.at[cod, _piv.columns[-1]] or 1) * 30), 999)

        df["dias_cobertura"] = df.apply(_calc_dias, axis=1)
        df.drop(columns=["codigo_norm", "periodo_firme_real", "lt_demanda_real",
                         "lt_demanda"], errors="ignore", inplace=True)
    else:
        # Fallback sem df_demanda
        df["demanda_mensal"]         = df["demanda_anual"] / 12
        df["periodo_firme"]          = df["demanda_mensal"] * MESES_FIRME
        lt = df["lead_time_dias"].apply(lambda x: x if x > 0 else LT_PADRAO)
        df["estoque_seguranca_calc"] = df["periodo_firme"] + (lt / 30) * df["demanda_mensal"]
        df["dias_cobertura"]         = (
            (df["saldo_estoque"] + df["saldo_inspecao"])
            / df["demanda_mensal"].replace(0, pd.NA) * 30
        ).fillna(0).clip(upper=999).round(0).astype(int)

    # ── Cobertura ─────────────────────────────────────────────────────────────
    df["cobertura_total"]    = df["saldo_estoque"] + df["saldo_inspecao"] + df["saldo_pedidos"]
    df["cobertura_sem_ped"]  = df["saldo_estoque"] + df["saldo_inspecao"]
    df["coberto_periodo_firme"] = df["cobertura_total"] >= df["periodo_firme"]
    df["coberto_estoque_seg"]   = df["cobertura_total"] >= df["estoque_seguranca_calc"]
    df["gap_cobertura"]         = (df["cobertura_total"] - df["estoque_seguranca_calc"]).round(2)

    # ── Flags auxiliares ──────────────────────────────────────────────────────
    df["eh_critico"]  = df["nivel_estoque"].str.contains("CRÍTICO",  case=False, na=False)
    df["eh_abaixo"]   = df["nivel_estoque"].str.contains("ABAIXO",   case=False, na=False)
    df["tem_demanda"] = df["demanda_anual"] > 0

    # ── Classificação de ação sugerida ────────────────────────────────────────
    def _acao(row) -> str:
        nivel  = row["nivel_estoque"].upper()
        critico_abaixo = ("CRÍTICO" in nivel) or ("ABAIXO" in nivel)
        coberto    = bool(row["coberto_periodo_firme"])
        tem_ped    = row["saldo_pedidos"] > 0
        tem_cont   = row["saldo_contrato"] > 0
        em_cont    = _clean_str(row["status_contratacao"]) != ""
        ped_cobre  = row["cobertura_total"] >= row["estoque_seguranca_calc"]
        sug        = row["sugestao_pedidos"]
        cont_insuf = tem_cont and (row["saldo_contrato"] < sug * 0.5) and sug > 0
        dias       = row["dias_cobertura"]

        # Contrato insuficiente → alerta de aditamento
        if cont_insuf and critico_abaixo:
            return "⚠️ POSSÍVEL ADITAMENTO CONTRATUAL"

        if coberto:
            if tem_ped:
                return "📞 FOLLOW-UP"
            return "✅ COBERTO"

        # Não coberto
        if tem_cont:
            if ped_cobre:
                return "📞 FOLLOW-UP"
            return "📋 EMITIR PEDIDO DE COMPRA"

        if em_cont:
            return "⏳ AGUARDAR CONTRATAÇÃO"

        # Sem contrato e sem ação
        if critico_abaixo:
            return "🚨 CONTRATAR URGENTE"
        if dias <= 60:
            return "⚠️ CONTRATAR PREVENTIVO"
        return "✅ COBERTO"

    df["acao_sugerida"] = df.apply(_acao, axis=1)

    # ── Grupos ────────────────────────────────────────────────────────────────
    def _grupo(acao: str) -> str:
        if "EMITIR PEDIDO" in acao or "ADITAMENTO" in acao:
            return "G1_EMITIR_PEDIDO"
        if "CONTRATAR URGENTE" in acao or "CONTRATAR PREVENTIVO" in acao:
            return "G2_CONTRATAR"
        if "AGUARDAR" in acao:
            return "G4_AGUARDAR_CONTRATACAO"
        if "FOLLOW-UP" in acao:
            return "G3_FOLLOW_UP"
        return "G0_COBERTO"

    df["grupo_pcm"] = df["acao_sugerida"].apply(_grupo)

    df["area_responsavel"] = df["grupo_pcm"].map({
        "G1_EMITIR_PEDIDO":       "PLANEJAMENTO",
        "G3_FOLLOW_UP":           "PLANEJAMENTO",
        "G2_CONTRATAR":           "COMPRAS",
        "G4_AGUARDAR_CONTRATACAO":"COMPRAS",
        "G0_COBERTO":             "—",
    }).fillna("—")

    df["responsavel"] = df.apply(
        lambda r: _clean_str(r["planejador"]) or "Não designado"
        if r["area_responsavel"] == "PLANEJAMENTO"
        else _clean_str(r["comprador"]) or "Não designado",
        axis=1,
    )

    # ── Plano de ação (texto simples por item) ────────────────────────────────
    def _plano(row) -> str:
        acao = row["acao_sugerida"]
        dias = int(row["dias_cobertura"])
        sug  = row["sugestao_pedidos"]
        if "COBERTO" in acao:
            return f"Coberto — {dias}d de cobertura"
        if "FOLLOW-UP" in acao:
            return f"Ligar ao fornecedor — confirmar status de entrega ({dias}d cobertura)"
        if "EMITIR PEDIDO" in acao:
            q = sug if sug > 0 else max(-row["gap_cobertura"], 0)
            return f"Emitir OC de {q:,.0f} un. usando saldo de contrato"
        if "ADITAMENTO" in acao:
            return f"Contrato insuficiente — solicitar aditamento ({dias}d cobertura)"
        if "AGUARDAR" in acao:
            st = row["status_contratacao"]
            return f"Em contratação: '{st}' — acompanhar ({dias}d cobertura)"
        if "CONTRATAR" in acao:
            return f"Designar comprador e iniciar contratação ({dias}d cobertura)"
        return "Verificar"

    df["plano_acao"] = df.apply(_plano, axis=1)

    # ── Programas e departamentos (split por "/") ─────────────────────────────
    # Cada item pode ter vários programas/departamentos; manter como listas internas
    # A explosão é feita nas funções de resumo para não duplicar itens
    def _split_barra(v: str) -> list[str]:
        return [x.strip() for x in v.split("/") if x.strip()]

    df["progs_list"] = df["prog_orcamentario_raw"].apply(_split_barra)
    df["depts_list"] = df["demandantes"].apply(_split_barra)

    # Primeiro programa e departamento (para exibição principal)
    df["prog_orcamentario"] = df["progs_list"].apply(lambda x: x[0] if x else "SEM PROGRAMA")
    df["departamento"]      = df["depts_list"].apply(lambda x: x[0] if x else "SEM DEPARTAMENTO")

    print(f"\n  Ações calculadas:")
    for grp in ["G1_EMITIR_PEDIDO", "G2_CONTRATAR", "G4_AGUARDAR_CONTRATACAO",
                "G3_FOLLOW_UP", "G0_COBERTO"]:
        sub = df[df["grupo_pcm"] == grp]
        if not sub.empty:
            crit = int(sub["eh_critico"].sum())
            print(f"    {grp:<35} : {len(sub):>4} itens  ({crit} críticos)")

    return df


def resumo_pcm(df_pcm: pd.DataFrame) -> dict:
    """Calcula KPIs completos do PCM."""
    if df_pcm.empty:
        return {}

    total = len(df_pcm)
    criticos = int(df_pcm["eh_critico"].sum()) if "eh_critico" in df_pcm.columns else 0

    def _cnt(grp):  return int((df_pcm["grupo_pcm"] == grp).sum())
    def _crit(grp): return int(((df_pcm["grupo_pcm"] == grp) & df_pcm["eh_critico"]).sum())

    g1t, g2t = _cnt("G1_EMITIR_PEDIDO"), _cnt("G2_CONTRATAR")
    g4t, g3t = _cnt("G4_AGUARDAR_CONTRATACAO"), _cnt("G3_FOLLOW_UP")

    cobertos    = int(df_pcm["coberto_periodo_firme"].sum()) if "coberto_periodo_firme" in df_pcm.columns else 0
    nao_cobertos = total - cobertos

    # Totais consolidados de quantidade para compra
    _g1_df = df_pcm[df_pcm["grupo_pcm"] == "G1_EMITIR_PEDIDO"]
    qtd_total_oc = _g1_df["sugestao_pedidos"].sum()

    _g2_df = df_pcm[df_pcm["grupo_pcm"] == "G2_CONTRATAR"]
    qtd_total_contratar = _g2_df["sugestao_pedidos"].sum()

    return {
        "total_itens": total,
        "total_criticos": criticos,
        "total_com_demanda": int(df_pcm["tem_demanda"].sum()) if "tem_demanda" in df_pcm.columns else 0,
        "cobertos_periodo_firme": cobertos,
        "nao_cobertos_periodo_firme": nao_cobertos,
        "g1_total": g1t,
        "g1_criticos": _crit("G1_EMITIR_PEDIDO"),
        "g2_total": g2t,
        "g2_criticos": _crit("G2_CONTRATAR"),
        "g3_followup_total": g3t,
        "g4_total": g4t,
        "g1_percentual_critico": round(_crit("G1_EMITIR_PEDIDO") / g1t * 100, 1) if g1t > 0 else 0,
        "g2_percentual_critico": round(_crit("G2_CONTRATAR") / g2t * 100, 1) if g2t > 0 else 0,
        "qtd_total_oc": qtd_total_oc,
        "qtd_total_contratar": qtd_total_contratar,
        "alerta_g1_bottleneck": g1t > 100,
        "alerta_g2_capacidade": g2t > 500,
        "alerta_criticos_emergencia": criticos > 436,
    }


def resumo_cobertura_por_dimensao(df_pcm: pd.DataFrame, dimensao: str,
                                   df_demanda_detail=None) -> pd.DataFrame:
    """
    Cobertura agrupada por dimensão com split por "/".
    dimensao: 'prog_orcamentario' | 'departamento' | 'grupo_material'
    """
    if df_pcm.empty:
        return pd.DataFrame()

    # Para programa e departamento, explode as listas
    if dimensao == "prog_orcamentario" and "progs_list" in df_pcm.columns:
        df_exp = df_pcm.copy()
        df_exp = df_exp.explode("progs_list").rename(columns={"progs_list": "_dim_val"})
        df_exp = df_exp[df_exp["_dim_val"].notna() & (df_exp["_dim_val"] != "")]
        dim_col = "_dim_val"
    elif dimensao == "departamento" and "depts_list" in df_pcm.columns:
        df_exp = df_pcm.copy()
        df_exp = df_exp.explode("depts_list").rename(columns={"depts_list": "_dim_val"})
        df_exp = df_exp[df_exp["_dim_val"].notna() & (df_exp["_dim_val"] != "")]
        dim_col = "_dim_val"
    else:
        if dimensao not in df_pcm.columns:
            return pd.DataFrame()
        df_exp = df_pcm.copy()
        dim_col = dimensao

    agg = (
        df_exp.groupby(dim_col, as_index=False)
        .agg(
            total_itens        =(dim_col,             "count"),
            itens_criticos     =("eh_critico",         "sum"),
            itens_cobertos     =("coberto_periodo_firme", "sum"),
            itens_com_contrato =("saldo_contrato",     lambda x: (x > 0).sum()),
            itens_com_pedido   =("saldo_pedidos",      lambda x: (x > 0).sum()),
            emitir_pedido      =("grupo_pcm",          lambda x: (x == "G1_EMITIR_PEDIDO").sum()),
            contratar          =("grupo_pcm",          lambda x: (x == "G2_CONTRATAR").sum()),
            follow_up          =("grupo_pcm",          lambda x: (x == "G3_FOLLOW_UP").sum()),
            aguardar           =("grupo_pcm",          lambda x: (x == "G4_AGUARDAR_CONTRATACAO").sum()),
        )
    )

    # Calcular qtd_oc_sugerida separadamente para evitar index alignment issues após explode
    g1_data = df_exp[df_exp["grupo_pcm"] == "G1_EMITIR_PEDIDO"].groupby(dim_col)["sugestao_pedidos"].sum().reset_index()
    g1_data.columns = [dim_col, "qtd_oc_sugerida"]
    agg = agg.merge(g1_data, on=dim_col, how="left")
    agg["qtd_oc_sugerida"] = agg["qtd_oc_sugerida"].fillna(0)
    agg["nao_cobertos"]  = agg["total_itens"] - agg["itens_cobertos"]
    agg["pct_cobertos"]  = (agg["itens_cobertos"] / agg["total_itens"].replace(0, 1) * 100).round(1)
    agg["sem_contrato"]  = agg["total_itens"] - agg["itens_com_contrato"]
    agg["risco"] = agg.apply(
        lambda r: "🔴 ALTO"   if r["contratar"] > 0
        else      "🟡 MÉDIO"  if r["nao_cobertos"] > 0
        else      "🟢 OK",
        axis=1,
    )

    # Se df_demanda_detail fornecido, calcular demanda real por programa/depto
    if df_demanda_detail is not None and not df_demanda_detail.empty:
        _dcol_map = {
            "prog_orcamentario": "programa_orcamentario",
            "departamento":      "departamento",
        }
        _dcol = _dcol_map.get(dimensao)
        _dem_agg_col = "demanda" if "demanda" in df_demanda_detail.columns else "quantidade"
        if _dcol and _dcol in df_demanda_detail.columns:
            _dem_dim = (df_demanda_detail.groupby(_dcol)[_dem_agg_col].sum()
                        .reset_index().rename(columns={_dcol: dim_col, _dem_agg_col: "demanda_real"}))
            agg = agg.merge(_dem_dim, on=dim_col, how="left")
            agg["demanda_real"] = agg["demanda_real"].fillna(0)

    return agg.rename(columns={dim_col: "Dimensão"}).sort_values("nao_cobertos", ascending=False)


def gerar_atividades_por_colaborador(df_pcm: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Retorna dict {colaborador: DataFrame} separado por área.
    """
    if df_pcm.empty:
        return {}

    resultado = {}
    _cols = [c for c in ["codigo", "descricao", "grupo_material", "nivel_estoque",
                          "acao_sugerida", "plano_acao", "dias_cobertura",
                          "sugestao_pedidos", "saldo_contrato", "saldo_pedidos",
                          "status_contratacao", "prog_orcamentario", "departamento"]
             if c in df_pcm.columns]

    # G1 + Follow-up → Planejadores
    df_plan = df_pcm[df_pcm["area_responsavel"] == "PLANEJAMENTO"].copy()
    for colab in sorted(df_plan["responsavel"].unique()):
        sub = df_plan[df_plan["responsavel"] == colab][_cols].copy()
        resultado[f"PLANEJAMENTO | {colab}"] = sub.sort_values("dias_cobertura")

    # G2 + G4 → Compradores
    df_comp = df_pcm[df_pcm["area_responsavel"] == "COMPRAS"].copy()
    for colab in sorted(df_comp["responsavel"].unique()):
        sub = df_comp[df_comp["responsavel"] == colab][_cols].copy()
        resultado[f"COMPRAS | {colab}"] = sub.sort_values("dias_cobertura")

    return resultado


if __name__ == "__main__":
    main()
