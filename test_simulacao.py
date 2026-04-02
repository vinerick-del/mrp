#!/usr/bin/env python3
"""
test_simulacao.py
=================
Simulação controlada do pipeline MRP com dados de exemplo em data_sim/.

Valida:
  1. Leitura de todos os arquivos de entrada
  2. Classificação ABC (A/B/C por valor acumulado)
  3. Cálculo MRP período a período (estoque projetado, SS, pedidos gerados)
  4. Janela de lead time — entradas_em_transito não mascara rupturas futuras
  5. Teto phase-out (horizonte_finito) — sem over-ordering além da demanda
  6. Classe C: trigger 1 mês, cobertura 4 meses
  7. Parcelas financeiras — datas, valores, nº da parcela, sem dobro na última
  8. Rateio por departamento/programa

Execute:  cd /home/user/mrp && python test_simulacao.py
"""
import os
import sys
import math
from datetime import date, timedelta
from typing import Any

import pandas as pd

# ── Setup ─────────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mrp

DIR_SIM    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_sim")
DIR_OUTPUT = os.path.join(DIR_SIM, "output")

# Redireciona salvar() e rateio_base.csv para data_sim
mrp.DIR_DADOS = DIR_SIM
mrp.DIR_SAIDA = DIR_OUTPUT
os.makedirs(DIR_OUTPUT, exist_ok=True)

from mrp import (
    transformar_demanda_dtm,
    ler_estoque_sap,
    ler_materiais,
    ler_lead_times,
    ler_politica_pagamento,
    passo_5_abc,
    passos_6_11_mrp,
    passo_12_rateio,
    br_to_float,
    _ler_sap_tabsep,
)

# ─────────────────────────────────────────────────────────────────────────────
# Utilidades de verificação
# ─────────────────────────────────────────────────────────────────────────────
_erros: list[str] = []
_ok:    list[str] = []

def check(label: str, actual: Any, expected: Any, tol: float = 0) -> None:
    if isinstance(expected, (int, float)) and tol > 0:
        ok = abs(float(actual) - float(expected)) <= tol
    else:
        ok = (actual == expected)
    if ok:
        _ok.append(label)
        print(f"  ✅ {label}: {actual}")
    else:
        _erros.append(label)
        print(f"  ❌ {label}  esperado={expected!r}  obtido={actual!r}")

def secao(titulo: str) -> None:
    print(f"\n{'═'*72}")
    print(f"  {titulo}")
    print(f"{'═'*72}")

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 1 — LEITURA DOS ARQUIVOS
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 1 │ LEITURA DOS ARQUIVOS DE ENTRADA")

# 1a. Demanda DTM
dem_raw = transformar_demanda_dtm(os.path.join(DIR_SIM, "Demanda.csv"))
dem_consolidada = (
    dem_raw.groupby(["material", "mes"], as_index=False)["quantidade"].sum()
)
check("Demanda: nº materiais", dem_raw["material"].nunique(), 4)
check("Demanda: total linhas", len(dem_raw), 48)   # 4 mats × 12 meses

# 1b. Estoque SAP
estoque = ler_estoque_sap(os.path.join(DIR_SIM, "ESTOQUE_SAP.csv"))

def est(mat):
    rows = estoque.loc[estoque["material"] == mat, "estoque_total"]
    return int(rows.iloc[0]) if len(rows) else None

check("Estoque 10001 = 3",   est("10001"),  3)
check("Estoque 10002 = 6",   est("10002"),  6)
check("Estoque 10003 = 25",  est("10003"), 25)
check("Estoque 10004 = 500", est("10004"), 500)

# 1c. Materiais
materiais = ler_materiais(os.path.join(DIR_SIM, "materiais.csv"))
check("Materiais: nº linhas", len(materiais), 4)
check("Preço 10001 = R$50.000",
      float(materiais.loc[materiais["material"] == "10001", "valor_unitario"].iloc[0]), 50000.0)
check("Preço 10004 = R$5",
      float(materiais.loc[materiais["material"] == "10004", "valor_unitario"].iloc[0]), 5.0)

# 1d. Lead Times
lt_dict = ler_lead_times(os.path.join(DIR_SIM, "LEAD_TIMES.csv"))
check("LT 10001 = 90 dias", lt_dict.get("10001"), 90)
check("LT 10002 = 60 dias", lt_dict.get("10002"), 60)
check("LT 10003 = 45 dias", lt_dict.get("10003"), 45)
check("LT 10004 = 7 dias",  lt_dict.get("10004"),  7)

# 1e. Pedidos Abertos — parser manual (replica passo_4_pedidos_abertos)
def ler_pedidos_direto(path: str):
    df = _ler_sap_tabsep(path)
    df.columns = df.columns.str.strip()
    col_mat = next(
        (c for c in df.columns if c.lower() in (
            "material", "nº material", "nr. material",
            "cod. material", "código material",
        )), df.columns[0]
    )
    df = df.rename(columns={col_mat: "material"})
    df["material"] = df["material"].astype(str).str.strip()

    col_qtd = next(
        (c for c in df.columns if c.lower() in (
            "quantidade", "qty", "qtd", "a ser fornecida (quantidade)",
        )), None
    )
    df["quantidade"] = df[col_qtd].apply(br_to_float) if col_qtd else 1.0

    col_val = next(
        (c for c in df.columns if c.lower() in (
            "a ser fornecido (valor)", "valor a ser fornecido",
            "valor total", "net value", "valor líquido total",
        )), None
    )
    df["valor_total_pedido"] = df[col_val].apply(br_to_float) if col_val else 0.0

    col_ent = next(
        (c for c in df.columns if c.lower() in (
            "data de remessa", "data remessa", "delivery date",
        )), None
    )
    col_cri = next(
        (c for c in df.columns if c.lower() in (
            "data do documento", "data doc.", "doc. date",
        )), None
    )
    df["data_remessa"] = (
        pd.to_datetime(df[col_ent], format="%d/%m/%Y", errors="coerce")
        if col_ent else pd.NaT
    )
    df["mes_remessa"] = df["data_remessa"].dt.to_period("M").astype(str)
    df["mes_pedido"]  = (
        pd.to_datetime(df[col_cri], format="%d/%m/%Y", errors="coerce")
        .dt.to_period("M").astype(str)
        if col_cri else None
    )
    col_num = next((c for c in df.columns if c.lower() in (
        "documento de compras", "nº doc. compras", "purchase order",
    )), None)
    df["numero_pedido"] = df[col_num].astype(str).str.strip() if col_num else None
    df["contrato"] = None
    df["fornecedor"] = "—"

    mes_atual = str(pd.Period(date.today(), "M"))
    atrasados = df["mes_remessa"] < mes_atual
    if atrasados.any():
        print(f"  ⚠ {atrasados.sum()} pedido(s) atrasado(s) → realocados para {mes_atual}")
        df.loc[atrasados, "mes_remessa"]  = mes_atual
        df.loc[atrasados, "data_remessa"] = pd.to_datetime(date.today())

    df_fut = df[df["mes_remessa"] >= mes_atual].copy()
    entradas = (
        df_fut.groupby(["material", "mes_remessa"], as_index=False)["quantidade"]
        .sum()
        .rename(columns={"mes_remessa": "mes", "quantidade": "qtd_entrada"})
    )
    return entradas, df_fut

ped_path = os.path.join(DIR_SIM, "pedidos_abertos.csv")
entradas_consolidadas, df_abertos = ler_pedidos_direto(ped_path)

check("Pedidos abertos: 3 linhas", len(df_abertos), 3)
check("PO 10001 qty = 4",
      int(df_abertos.loc[df_abertos["material"] == "10001", "quantidade"].sum()), 4)
check("PO 10001 valor = R$200.000",
      float(df_abertos.loc[df_abertos["material"] == "10001", "valor_total_pedido"].sum()), 200000.0)
check("PO 10002 qty = 8",
      int(df_abertos.loc[df_abertos["material"] == "10002", "quantidade"].sum()), 8)
check("PO 10003 qty = 30",
      int(df_abertos.loc[df_abertos["material"] == "10003", "quantidade"].sum()), 30)

# 1f. Política de pagamento
pol_path = os.path.join(DIR_SIM, "Politica_de_pagamento")
politica = ler_politica_pagamento(pol_path)
check("Política 4500100001 = [30, 60]",    politica.get("4500100001"), [30, 60])
check("Política 4500100002 = [30]",        politica.get("4500100002"), [30])
check("Política 4500100003 = [30, 60, 90]",politica.get("4500100003"), [30, 60, 90])

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 2 — CLASSIFICAÇÃO ABC
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 2 │ CLASSIFICAÇÃO ABC")
# Valores anuais esperados:
#   10001: 12 un × R$50.000  =  R$600.000  →  56,5% acum → A
#   10002: 24 un × R$15.000  =  R$360.000  →  90,4% acum → B
#   10003: 120 un × R$800    =   R$96.000  →  99,4% acum → C
#   10004: 1200 un × R$5     =    R$6.000  → 100,0% acum → C
#   Total = R$1.062.000

abc = passo_5_abc(dem_consolidada, materiais)

def abc_row(mat):
    rows = abc[abc["material"] == mat]
    return rows.iloc[0] if len(rows) else None

r1, r2, r3, r4 = abc_row("10001"), abc_row("10002"), abc_row("10003"), abc_row("10004")

check("ABC 10001 = A", r1["classe"], "A")
check("ABC 10002 = B", r2["classe"], "B")
check("ABC 10003 = C", r3["classe"], "C")
check("ABC 10004 = C", r4["classe"], "C")
check("Valor total 10001 = R$600.000", float(r1["valor_total"]),  600000.0, tol=1)
check("Valor total 10002 = R$360.000", float(r2["valor_total"]),  360000.0, tol=1)
check("Valor total 10003 = R$96.000",  float(r3["valor_total"]),   96000.0, tol=1)
check("Valor total 10004 = R$6.000",   float(r4["valor_total"]),    6000.0, tol=1)

print("\n  Tabela ABC:")
print(abc[["material", "classe", "valor_total", "pct_acum"]].to_string(index=False))

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 3 — MRP MÊS A MÊS
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 3 │ MRP — MOTOR PRINCIPAL")

df_mrp, df_ped = passos_6_11_mrp(
    demanda          = dem_consolidada,
    estoque          = estoque,
    entradas_pedidos = entradas_consolidadas,
    abc              = abc,
    materiais        = materiais,
    lead_times_dict  = lt_dict,
    horizonte_finito = True,
)

def mrp_row(mat, per):
    rows = df_mrp[(df_mrp["material"] == mat) & (df_mrp["periodo"] == per)]
    return rows.iloc[0] if len(rows) else None

def ep(mat, per):
    r = mrp_row(mat, per)
    return int(r["estoque_projetado"]) if r is not None else None

def ped_gen(mat, per):
    r = mrp_row(mat, per)
    return int(r["pedido_gerado"]) if r is not None else None

# ── 3a. Material 10001 (Classe A, LT=90d → 3 meses) ─────────────────────────
print("\n  [10001 — Classe A, LT=90d]")
print(df_mrp[df_mrp["material"] == "10001"][
    ["periodo","demanda","total_entradas","estoque_projetado",
     "estoque_seguranca_3m","pedido_gerado","periodo_entrega"]
].to_string(index=False))

# Open order de 4 un chega em 2026-05
check("10001 est_proj 2026-04 = 2", ep("10001", "2026-04"),  2)
check("10001 est_proj 2026-05 = 5 (PO chega)", ep("10001", "2026-05"),  5)
check("10001 est_proj 2026-06 = 4", ep("10001", "2026-06"),  4)
check("10001 est_proj 2026-07 = 3", ep("10001", "2026-07"),  3)
check("10001 est_proj 2026-08 = 2", ep("10001", "2026-08"),  2)
check("10001 est_proj 2026-09 = 1", ep("10001", "2026-09"),  1)
# Nenhum pedido até ago (janela LT cobre o PO aberto)
check("10001 sem pedido 2026-04 (LT window cobre PO)", ped_gen("10001", "2026-04"), 0)

ped_10001 = df_ped[df_ped["material"] == "10001"].sort_values("periodo_necessidade")
check("10001: nº pedidos gerados = 3", len(ped_10001), 3)
if len(ped_10001) >= 1:
    check("10001: 1º necessidade em 2026-08", ped_10001.iloc[0]["periodo_necessidade"], "2026-08")
    check("10001: 1º entrega em 2026-10",     ped_10001.iloc[0]["periodo_entrega"],      "2026-10")
    check("10001: 1º qty = 1",                int(ped_10001.iloc[0]["quantidade"]),       1)

# ── 3b. Material 10002 (Classe B, LT=60d → 2 meses) ─────────────────────────
print("\n  [10002 — Classe B, LT=60d]")
print(df_mrp[df_mrp["material"] == "10002"][
    ["periodo","demanda","total_entradas","estoque_projetado",
     "estoque_seguranca_3m","pedido_gerado","periodo_entrega"]
].to_string(index=False))

# Open order 8 un chega 2026-06; LT=2m → em 2026-04, idx_chegada=2 → abrange jun
check("10002 est_proj 2026-06 = 8 (PO chega)", ep("10002", "2026-06"),  8)
check("10002 sem pedido 2026-04 (LT window cobre PO)", ped_gen("10002", "2026-04"), 0)

ped_10002 = df_ped[df_ped["material"] == "10002"].sort_values("periodo_necessidade")
check("10002: nº pedidos gerados = 3", len(ped_10002), 3)
if len(ped_10002) >= 1:
    check("10002: 1º necessidade em 2026-08", ped_10002.iloc[0]["periodo_necessidade"], "2026-08")
    check("10002: 1º entrega em 2026-09",     ped_10002.iloc[0]["periodo_entrega"],      "2026-09")
    check("10002: 1º qty = 2",                int(ped_10002.iloc[0]["quantidade"]),       2)

# ── 3c. Material 10003 (Classe C, LT=45d → 2 meses) ─────────────────────────
print("\n  [10003 — Classe C, LT=45d, Cobertura=4m]")
print(df_mrp[df_mrp["material"] == "10003"][
    ["periodo","demanda","total_entradas","estoque_projetado",
     "estoque_seguranca_3m","pedido_gerado","periodo_entrega"]
].to_string(index=False))

# Trigger C: estoque_virtual < dem do mês → 1ª trigger em 2026-08 (est=5 < dem=10)
check("10003: nenhum pedido antes de ago (estoque > dem)",
      all(ped_gen("10003", p) == 0 for p in ["2026-04","2026-05","2026-06","2026-07"]), True)

ped_10003 = df_ped[df_ped["material"] == "10003"].sort_values("periodo_necessidade")
check("10003: nº pedidos gerados = 2", len(ped_10003), 2)
if len(ped_10003) >= 1:
    check("10003: 1º trigger em 2026-08", ped_10003.iloc[0]["periodo_necessidade"], "2026-08")
    check("10003: 1º entrega em 2026-09", ped_10003.iloc[0]["periodo_entrega"],      "2026-09")
    # qty: ss_ordem(40) - est_virtual(5) = 35  →  teto permite (50-5=45)
    check("10003: 1º qty = 35 (cobertura 4m)", int(ped_10003.iloc[0]["quantidade"]), 35)

# ── 3d. Material 10004 (Classe C, LT=7d → 1 mês, bump para mês seguinte) ────
print("\n  [10004 — Classe C, LT=7d, entrega bumped +1m]")
print(df_mrp[df_mrp["material"] == "10004"][
    ["periodo","demanda","total_entradas","estoque_projetado",
     "estoque_seguranca_3m","pedido_gerado","periodo_entrega"]
].to_string(index=False))

ped_10004 = df_ped[df_ped["material"] == "10004"].sort_values("periodo_necessidade")
check("10004: nº pedidos gerados = 2", len(ped_10004), 2)
if len(ped_10004) >= 1:
    check("10004: 1º trigger em 2026-08", ped_10004.iloc[0]["periodo_necessidade"], "2026-08")
    # LT=7d: 2026-08-01 + 7 = 2026-08-08 → mesmo mês → bump para 2026-09
    check("10004: 1º entrega = 2026-09 (LT<1m → bump)",
          ped_10004.iloc[0]["periodo_entrega"], "2026-09")
    check("10004: 1º qty = 400 (cobertura 4m×100)", int(ped_10004.iloc[0]["quantidade"]), 400)

# ── 3e. Estoque nunca negativo em nenhum material ────────────────────────────
for mat in ["10001", "10002", "10003", "10004"]:
    min_est = int(df_mrp[df_mrp["material"] == mat]["estoque_projetado"].min())
    check(f"{mat}: estoque >= 0 em todo horizonte (mín={min_est})", min_est >= 0, True)

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 4 — TETO PHASE-OUT (horizonte_finito)
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 4 │ TETO PHASE-OUT — horizonte_finito=True")

# Após idx_fim_dem (2026-12), não deve haver pedidos com entrega depois de 2027-01
# A única exceção seria se o teto permitisse — mas com demanda = 0 após dez/2026,
# não há necessidade de compra para entregar em fev/2027 ou além.
excesso = df_ped[
    (df_ped["periodo_entrega"] != "além_horizonte") &
    (df_ped["periodo_entrega"] > "2027-01")
]
check("Nenhum pedido com entrega além de 2027-01", len(excesso), 0)
if not excesso.empty:
    print("  Pedidos excessivos encontrados:")
    print(excesso[["material","periodo_necessidade","periodo_entrega","quantidade"]].to_string(index=False))

# Valor total gerado vs. demanda restante (sanity check)
val_gerado = df_ped["valor_total_pedido"].sum() if not df_ped.empty else 0
print(f"\n  Valor total pedidos GERADOS: R$ {val_gerado:,.2f}")
print(f"  Valor total pedidos ABERTOS: R$ {df_abertos['valor_total_pedido'].sum():,.2f}")

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 5 — PARCELAS FINANCEIRAS
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 5 │ PARCELAS FINANCEIRAS (PEDIDOS ABERTOS)")

def calcular_parcelas(df_abertos_fut: pd.DataFrame, politica_pag: dict) -> pd.DataFrame:
    """Replica a lógica de app.py para gerar o fluxo de parcelas."""
    parcelas = []
    for _, row in df_abertos_fut.iterrows():
        data_base    = row["data_remessa"]
        contrato_ref = row.get("contrato")
        pedido_ref   = row.get("numero_pedido")
        dias         = None
        if contrato_ref and str(contrato_ref) not in ("None", "nan", ""):
            dias = politica_pag.get(str(contrato_ref))
        if dias is None and pedido_ref and str(pedido_ref) not in ("None", "nan", ""):
            dias = politica_pag.get(str(pedido_ref))
        if dias is None:
            dias = [60, 90]   # default app.py

        n     = len(dias)
        vbase = round(row["valor_total_pedido"] / n, 2)
        resto = round(row["valor_total_pedido"] - vbase * (n - 1), 2)

        for i, d in enumerate(dias):
            parcelas.append({
                "material"      : row["material"],
                "numero_pedido" : str(pedido_ref),
                "prazo_dias"    : d,
                "mes_pagamento" : (data_base + timedelta(days=d)).strftime("%Y-%m"),
                "valor_parcela" : resto if i == n - 1 else vbase,
                "num_parcela"   : i + 1,
                "tot_parcelas"  : n,
                "valor_pedido"  : row["valor_total_pedido"],
            })
    return pd.DataFrame(parcelas)

df_parcelas = calcular_parcelas(df_abertos, politica)

print("\n  Parcelas geradas:")
print(df_parcelas[["numero_pedido","num_parcela","tot_parcelas",
                    "prazo_dias","mes_pagamento","valor_pedido",
                    "valor_parcela"]].to_string(index=False))

# ── PO 4500100001: R$200.000 / 2 parcelas iguais (30/60 dias) ────────────────
# data_base = 01/05/2026
# parcela 1: 100.000 → +30d = 31/05 → 2026-05
# parcela 2: 100.000 → +60d = 30/06 → 2026-06
po1 = df_parcelas[df_parcelas["numero_pedido"] == "4500100001"].sort_values("num_parcela")
check("PO1 nº parcelas = 2",         len(po1), 2)
check("PO1 parc1 mês = 2026-05",     po1.iloc[0]["mes_pagamento"], "2026-05")
check("PO1 parc1 valor = R$100.000", float(po1.iloc[0]["valor_parcela"]), 100000.0, tol=1)
check("PO1 parc2 mês = 2026-06",     po1.iloc[1]["mes_pagamento"], "2026-06")
check("PO1 parc2 valor = R$100.000", float(po1.iloc[1]["valor_parcela"]), 100000.0, tol=1)
check("PO1 soma parcelas = valor total",
      float(po1["valor_parcela"].sum()), float(po1.iloc[0]["valor_pedido"]), tol=0.01)

# ── PO 4500100002: R$120.000 / 1 parcela (30 dias) ───────────────────────────
# data_base = 01/06/2026
# parcela 1: 120.000 → +30d = 01/07 → 2026-07
po2 = df_parcelas[df_parcelas["numero_pedido"] == "4500100002"]
check("PO2 nº parcelas = 1",         len(po2), 1)
check("PO2 parc1 mês = 2026-07",     po2.iloc[0]["mes_pagamento"], "2026-07")
check("PO2 parc1 valor = R$120.000", float(po2.iloc[0]["valor_parcela"]), 120000.0, tol=1)

# ── PO 4500100003: R$24.000 / 3 parcelas iguais (30/60/90 dias) ─────────────
# data_base = 01/05/2026
# vbase = 24000/3 = 8000; _resto = 24000 - 8000*2 = 8000 (todas iguais)
# parcela 1: 8.000 → +30d = 31/05 → 2026-05
# parcela 2: 8.000 → +60d = 30/06 → 2026-06
# parcela 3: 8.000 → +90d = 30/07 → 2026-07
po3 = df_parcelas[df_parcelas["numero_pedido"] == "4500100003"].sort_values("num_parcela")
check("PO3 nº parcelas = 3",                  len(po3), 3)
check("PO3 parc1 mês = 2026-05",              po3.iloc[0]["mes_pagamento"], "2026-05")
check("PO3 parc2 mês = 2026-06",              po3.iloc[1]["mes_pagamento"], "2026-06")
check("PO3 parc3 mês = 2026-07",              po3.iloc[2]["mes_pagamento"], "2026-07")
check("PO3 parc1 valor = R$8.000",            float(po3.iloc[0]["valor_parcela"]), 8000.0, tol=1)
check("PO3 última parcela = _resto (R$8.000)",float(po3.iloc[2]["valor_parcela"]), 8000.0, tol=1)
check("PO3 soma parcelas = R$24.000",         float(po3["valor_parcela"].sum()), 24000.0, tol=0.01)
# Bug fix: última parcela NÃO é _vbase + _resto (seria 8000+8000=16000)
check("PO3 última parcela != dobro (bug fix corrigido)",
      float(po3.iloc[2]["valor_parcela"]) < 16000, True)

# ── Fluxo de caixa consolidado ────────────────────────────────────────────────
fluxo = df_parcelas.groupby("mes_pagamento")["valor_parcela"].sum().sort_index()
print("\n  Fluxo de caixa (pedidos abertos):")
for mes, val in fluxo.items():
    print(f"    {mes}: R$ {val:>12,.2f}")

# 2026-05: PO1 parc1 (100k) + PO3 parc1 (8k)          = 108.000
# 2026-06: PO1 parc2 (100k) + PO3 parc2 (8k)          = 108.000
# 2026-07: PO2 parc1 (120k) + PO3 parc3 (8k)          = 128.000
check("Fluxo 2026-05 = R$108.000", float(fluxo.get("2026-05", 0)), 108000.0, tol=1)
check("Fluxo 2026-06 = R$108.000", float(fluxo.get("2026-06", 0)), 108000.0, tol=1)
check("Fluxo 2026-07 = R$128.000", float(fluxo.get("2026-07", 0)), 128000.0, tol=1)
check("Soma fluxo = total POs (R$344.000)",
      float(fluxo.sum()), 344000.0, tol=0.01)

# ═══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 6 — RATEIO
# ═══════════════════════════════════════════════════════════════════════════════
secao("SEÇÃO 6 │ RATEIO POR DEPARTAMENTO / PROGRAMA ORÇAMENTÁRIO")

df_rateio = passo_12_rateio(df_ped, df_abertos)

if df_rateio.empty:
    print("  ⚠ Rateio retornou DataFrame vazio.")
else:
    cols = [c for c in ["material","departamento","programa_orcamentario",
                         "qtd_rateada","proporcao_pct"] if c in df_rateio.columns]
    print(df_rateio[cols].drop_duplicates().sort_values(
        ["material","departamento"]).to_string(index=False))

    # 10001: 100% DPC / A40_INSPECAO
    r10001 = df_rateio[df_rateio["material"] == "10001"][
        ["departamento","programa_orcamentario","proporcao_pct"]
    ].drop_duplicates()
    check("10001 rateio: 1 destino único", len(r10001), 1)
    check("10001 rateio: 100% A40_INSPECAO",
          float(r10001.iloc[0]["proporcao_pct"]), 100.0, tol=0.01)

    # 10002: 70% A40 + 30% A50
    r10002 = df_rateio[df_rateio["material"] == "10002"][
        ["departamento","programa_orcamentario","proporcao_pct"]
    ].drop_duplicates()
    check("10002 rateio: 2 destinos", len(r10002), 2)
    p_a40 = r10002.loc[r10002["programa_orcamentario"] == "A40_INSPECAO", "proporcao_pct"]
    p_a50 = r10002.loc[r10002["programa_orcamentario"] == "A50_MANUTENCAO", "proporcao_pct"]
    if not p_a40.empty:
        check("10002 rateio A40 = 70%", float(p_a40.iloc[0]), 70.0, tol=0.01)
    if not p_a50.empty:
        check("10002 rateio A50 = 30%", float(p_a50.iloc[0]), 30.0, tol=0.01)

    # 10003: 40% A40 + 60% A50
    r10003 = df_rateio[df_rateio["material"] == "10003"][
        ["departamento","programa_orcamentario","proporcao_pct"]
    ].drop_duplicates()
    check("10003 rateio: 2 destinos", len(r10003), 2)
    p3_a40 = r10003.loc[r10003["programa_orcamentario"] == "A40_INSPECAO", "proporcao_pct"]
    if not p3_a40.empty:
        check("10003 rateio A40 = 40%", float(p3_a40.iloc[0]), 40.0, tol=0.01)

# ═══════════════════════════════════════════════════════════════════════════════
# RESULTADO FINAL
# ═══════════════════════════════════════════════════════════════════════════════
secao("RESULTADO FINAL")
total = len(_ok) + len(_erros)
print(f"\n  ✅ Passou : {len(_ok)}/{total}")
print(f"  ❌ Falhou : {len(_erros)}/{total}")
if _erros:
    print("\n  Verificações com FALHA:")
    for e in _erros:
        print(f"    • {e}")
else:
    print("\n  🎉 Todas as verificações passaram!")
print()
