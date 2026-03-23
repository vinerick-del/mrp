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

import math
import os
from datetime import date

import numpy as np
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────
MESES_COBERTURA_SS     = 3   # Cobertura do estoque de segurança (A/B) — rolling
HORIZONTE_MESES        = 12  # Horizonte de planejamento
LEAD_TIME_MESES        = 1   # Lead time padrão

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
# PASSO 1-2: LER E CONSOLIDAR DEMANDA
# ─────────────────────────────────────────────────────────────────────────────
def passo_1_2_demanda() -> pd.DataFrame:
    separador("PASSO 1-2 │ LER E CONSOLIDAR DEMANDA")

    df = pd.read_csv(os.path.join(DIR_DADOS, "demanda.csv"))
    df["mes"] = pd.to_datetime(df["mes"], format="%Y-%m").dt.to_period("M").astype(str)

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

    df = pd.read_csv(os.path.join(DIR_DADOS, "estoque.csv"))
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
def passo_5_abc(demanda: pd.DataFrame, materiais: pd.DataFrame) -> pd.DataFrame:
    separador("PASSO 5 │ CLASSIFICAÇÃO ABC")

    dem_total = (
        demanda.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "demanda_total"})
    )
    abc = dem_total.merge(materiais[["material", "valor_unitario"]], on="material", how="left")
    abc["valor_unitario"] = abc["valor_unitario"].fillna(0)
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
) -> tuple[pd.DataFrame, pd.DataFrame]:
    separador("PASSOS 6-11 │ CÁLCULO MRP MÊS A MÊS")

    hoje      = date.today()
    per_atual = pd.Period(hoje, "M")
    periodos  = [str(p) for p in pd.period_range(start=per_atual, periods=HORIZONTE_MESES, freq="M")]
    n_per     = len(periodos)

    print(f"  Data atual  : {hoje.strftime('%d/%m/%Y')}")
    print(f"  Horizonte   : {periodos[0]} → {periodos[-1]}  ({n_per} meses)")

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
                pedido  = math.ceil(nec)
                idx_ent = i + LEAD_TIME_MESES

                if idx_ent < n_per:
                    per_entrega = periodos[idx_ent]
                    novas_ent[per_entrega] += pedido
                else:
                    per_entrega = "além_horizonte"

                pedidos_compra.append(
                    {
                        "material"           : mat,
                        "descricao"          : desc_map.get(mat, ""),
                        "classe"             : classe,
                        "periodo_necessidade": per,
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
            columns=["material", "descricao", "classe",
                     "periodo_necessidade", "periodo_entrega", "quantidade"]
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
) -> pd.DataFrame:
    separador("PASSO 12 │ RATEIO POR DEPARTAMENTO / PROGRAMA ORÇAMENTÁRIO")

    rateio_base = pd.read_csv(os.path.join(DIR_DADOS, "rateio_base.csv"))

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
    print(f"  Horizonte         : {HORIZONTE_MESES} meses  │  Lead time: {LEAD_TIME_MESES} mês(es)")
    print(f"  SS classes A/B    : {MESES_COBERTURA_SS} meses rolling")
    print(f"  Classe C          : trigger <1m cobertura → pedido cobre {CLASSE_C_COBERTURA_MESES}m (sem limite/ano)")

    os.makedirs(DIR_DADOS, exist_ok=True)
    os.makedirs(DIR_SAIDA, exist_ok=True)

    materiais                 = pd.read_csv(os.path.join(DIR_DADOS, "materiais.csv"))
    demanda                   = passo_1_2_demanda()
    estoque                   = passo_3_estoque()
    entradas, df_abertos_fut  = passo_4_pedidos_abertos()
    abc                       = passo_5_abc(demanda, materiais)
    df_mrp, df_ped            = passos_6_11_mrp(demanda, estoque, entradas, abc, materiais)
    df_rateio                 = passo_12_rateio(df_ped, df_abertos_fut)

    _imprimir_mrp_pivot(df_mrp)

    # Validações obrigatórias
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
