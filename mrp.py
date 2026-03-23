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
  11. Aplicar restrições por classe
  12. Gerar rateio final
"""

import math
import os
from datetime import date

import numpy as np
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────
MESES_COBERTURA_SS = 3        # Meses de cobertura do estoque de segurança (rolling)
HORIZONTE_MESES    = 12       # Meses de horizonte de planejamento
LEAD_TIME_MESES    = 1        # Lead time padrão em meses

LIMITE_ABC_A = 0.80           # Classe A → até 80% do valor acumulado
LIMITE_ABC_B = 0.95           # Classe B → 80% a 95%
MAX_PEDIDOS_CLASSE_C = 2      # Classe C → máximo de pedidos gerados por ano

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

    # Consolida: múltiplas linhas do mesmo material/mês são somadas
    demanda = df.groupby(["material", "mes"], as_index=False)["quantidade"].sum()

    print(f"  Materiais com demanda : {demanda['material'].nunique()}")
    print(f"  Período da demanda    : {demanda['mes'].min()} → {demanda['mes'].max()}")
    print(f"  Registros consolidados: {len(demanda)}")
    return demanda


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3: CONSOLIDAR ESTOQUE POR MATERIAL (IGNORAR ENDEREÇAMENTO)
# ─────────────────────────────────────────────────────────────────────────────
def passo_3_estoque() -> tuple[pd.DataFrame, pd.DataFrame]:
    separador("PASSO 3 │ CONSOLIDAR ESTOQUE (IGNORAR ENDEREÇAMENTO)")

    df_detalhe = pd.read_csv(os.path.join(DIR_DADOS, "estoque.csv"))
    print(f"  Linhas de endereçamento: {len(df_detalhe)}")

    # Regra: somar todas as quantidades por material, ignorar o endereço
    consolidado = (
        df_detalhe.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "estoque_total"})
    )

    print(f"  Materiais únicos      : {len(consolidado)}")
    print()
    print(consolidado.to_string(index=False))

    salvar(consolidado, "01_estoque_consolidado.csv")
    return consolidado, df_detalhe  # df_detalhe usado no rateio


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 4: LER PEDIDOS EM ABERTO (ENTRADAS FUTURAS)
# ─────────────────────────────────────────────────────────────────────────────
def passo_4_pedidos_abertos() -> pd.DataFrame:
    separador("PASSO 4 │ PEDIDOS EM ABERTO — ENTRADAS FUTURAS")

    df = pd.read_csv(
        os.path.join(DIR_DADOS, "pedidos_abertos.csv"),
        parse_dates=["data_remessa"],
    )

    # REGRA CRÍTICA: usar EXCLUSIVAMENTE data_remessa (NÃO data do pedido)
    df["mes_remessa"] = df["data_remessa"].dt.to_period("M").astype(str)

    hoje      = date.today()
    mes_atual = str(pd.Period(hoje, "M"))

    # Ignorar remessas já vencidas (meses passados)
    df_fut = df[df["mes_remessa"] >= mes_atual].copy()

    print(f"  Pedidos em aberto     : {len(df)}")
    print(f"  Mês de referência     : {mes_atual}")
    print(f"  Remessas futuras      : {len(df_fut)}")
    print()
    if not df_fut.empty:
        print(df_fut[["numero_pedido", "material", "quantidade", "data_remessa"]].to_string(index=False))

    # Consolidar por material e mês de remessa
    entradas = (
        df_fut.groupby(["material", "mes_remessa"], as_index=False)["quantidade"]
        .sum()
        .rename(columns={"mes_remessa": "mes", "quantidade": "qtd_entrada"})
    )
    return entradas


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 5: CLASSIFICAÇÃO ABC
# ─────────────────────────────────────────────────────────────────────────────
def passo_5_abc(demanda: pd.DataFrame, materiais: pd.DataFrame) -> pd.DataFrame:
    separador("PASSO 5 │ CLASSIFICAÇÃO ABC")

    # Base: valor_total = demanda_total × valor_unitário
    dem_total = (
        demanda.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "demanda_total"})
    )
    abc = dem_total.merge(materiais[["material", "valor_unitario"]], on="material", how="left")
    abc["valor_unitario"] = abc["valor_unitario"].fillna(0)
    abc["valor_total"] = abc["demanda_total"] * abc["valor_unitario"]

    # Ordenar do maior para o menor valor
    abc = abc.sort_values("valor_total", ascending=False).reset_index(drop=True)

    # Percentual individual e acumulado
    total_geral       = abc["valor_total"].sum()
    abc["pct_ind"]    = abc["valor_total"] / total_geral * 100
    abc["pct_acum"]   = abc["valor_total"].cumsum() / total_geral

    # Classificação
    abc["classe"] = "C"
    abc.loc[abc["pct_acum"] <= LIMITE_ABC_B, "classe"] = "B"
    abc.loc[abc["pct_acum"] <= LIMITE_ABC_A, "classe"] = "A"

    # Exibição
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
    contagem = abc.groupby("classe").size()
    print(
        f"\n  Classe A (≤{LIMITE_ABC_A*100:.0f}%): {contagem.get('A', 0)} material(is)"
        f"  │  Classe B ({LIMITE_ABC_A*100:.0f}%-{LIMITE_ABC_B*100:.0f}%): {contagem.get('B', 0)} material(is)"
        f"  │  Classe C (>{LIMITE_ABC_B*100:.0f}%): {contagem.get('C', 0)} material(is)"
    )

    # Exportar
    export = abc.copy()
    export["pct_ind"]  = export["pct_ind"].round(2)
    export["pct_acum"] = (export["pct_acum"] * 100).round(2)
    export = export.rename(columns={"pct_ind": "pct_individual", "pct_acum": "pct_acumulado"})
    salvar(export, "02_classificacao_abc.csv")

    return abc  # retorna com pct_acum como fração (0-1) para uso interno


# ─────────────────────────────────────────────────────────────────────────────
# PASSOS 6-11: CÁLCULO MRP MÊS A MÊS
#
#   6.  Iniciar cálculo MRP mês a mês
#   7.  Inserir entradas (pedidos existentes + novos)
#   8.  Calcular estoque projetado
#   9.  Calcular estoque de segurança e necessidade
#   10. Gerar compras
#   11. Aplicar restrições por classe
# ─────────────────────────────────────────────────────────────────────────────
def passos_6_11_mrp(
    demanda: pd.DataFrame,
    estoque: pd.DataFrame,
    entradas_pedidos: pd.DataFrame,
    abc: pd.DataFrame,
    materiais: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    separador("PASSOS 6-11 │ CÁLCULO MRP MÊS A MÊS")

    # ── Horizonte de planejamento ─────────────────────────────────────────────
    hoje        = date.today()
    per_atual   = pd.Period(hoje, "M")
    periodos    = [str(p) for p in pd.period_range(start=per_atual, periods=HORIZONTE_MESES, freq="M")]
    n_per       = len(periodos)

    print(f"  Data atual  : {hoje.strftime('%d/%m/%Y')}")
    print(f"  Horizonte   : {periodos[0]} → {periodos[-1]}  ({n_per} meses)")

    # ── Lookups ───────────────────────────────────────────────────────────────
    # demanda: {material: {periodo: qtd}}
    dem_lkp: dict[str, dict[str, float]] = {}
    for _, row in demanda.iterrows():
        dem_lkp.setdefault(row["material"], {})[row["mes"]] = float(row["quantidade"])

    # entradas de pedidos existentes: {material: {periodo: qtd}}
    ent_lkp: dict[str, dict[str, float]] = {}
    for _, row in entradas_pedidos.iterrows():
        ent_lkp.setdefault(row["material"], {})[row["mes"]] = float(row["qtd_entrada"])

    classe_map  = dict(zip(abc["material"], abc["classe"]))
    estoque_map = dict(zip(estoque["material"], estoque["estoque_total"]))
    desc_map    = dict(zip(materiais["material"], materiais["descricao"]))

    # Todos os materiais: união de estoque + demanda
    todos_mats = sorted(set(estoque["material"]) | set(demanda["material"]))

    resultados: list[dict] = []
    pedidos_compra: list[dict] = []

    # ── Loop por material ─────────────────────────────────────────────────────
    for mat in todos_mats:
        classe    = classe_map.get(mat, "C")
        est_ini   = float(estoque_map.get(mat, 0))
        dem_mat   = dem_lkp.get(mat, {})
        ent_mat   = ent_lkp.get(mat, {})

        # Entradas de novos pedidos gerados durante o loop (alimentadas iterativamente)
        novas_ent: dict[str, float] = {p: 0.0 for p in periodos}

        pedidos_c = 0        # Contador de pedidos gerados para Classe C
        est_proj  = est_ini  # Estoque projetado acumulado (rolling)

        for i, per in enumerate(periodos):
            dem = dem_mat.get(per, 0.0)

            # ── PASSO 9 │ Estoque de Segurança (rolling 3 meses) ─────────────
            # Cobertura = mês atual + próximos (MESES_COBERTURA_SS-1) meses
            ss = sum(
                dem_mat.get(periodos[j], 0.0)
                for j in range(i, min(i + MESES_COBERTURA_SS, n_per))
            )

            # ── PASSO 7 │ Entradas deste mês ─────────────────────────────────
            ent_exist = ent_mat.get(per, 0.0)
            ent_nova  = novas_ent.get(per, 0.0)
            total_ent = ent_exist + ent_nova

            # ── PASSO 8 │ Estoque Projetado ───────────────────────────────────
            est_proj = est_proj + total_ent - dem

            # ── PASSO 9 │ Necessidade ─────────────────────────────────────────
            nec = max(0.0, ss - est_proj)

            # ── PASSOS 10-11 │ Gerar compra com restrições de classe ──────────
            pedido      = 0
            per_entrega = ""
            bloq_c      = False

            if nec > 0:
                # Restrição Classe C: máximo MAX_PEDIDOS_CLASSE_C por ano
                if classe == "C" and pedidos_c >= MAX_PEDIDOS_CLASSE_C:
                    bloq_c = True
                else:
                    # Quantidade exata (arredondada para cima)
                    pedido = math.ceil(nec)

                    # Pedido chega em M + lead_time
                    idx_ent = i + LEAD_TIME_MESES
                    if idx_ent < n_per:
                        per_entrega = periodos[idx_ent]
                        novas_ent[per_entrega] += pedido
                    else:
                        per_entrega = "além_horizonte"

                    if classe == "C":
                        pedidos_c += 1

                    pedidos_compra.append(
                        {
                            "material"          : mat,
                            "descricao"         : desc_map.get(mat, ""),
                            "classe"            : classe,
                            "periodo_necessidade": per,
                            "periodo_entrega"   : per_entrega,
                            "quantidade"        : pedido,
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
                    "estoque_seguranca"         : int(ss),
                    "necessidade"               : int(nec),
                    "pedido_gerado"             : pedido,
                    "periodo_entrega"           : per_entrega,
                    "bloqueado_classe_c"        : bloq_c,
                }
            )

    df_mrp = pd.DataFrame(resultados)
    df_ped = (
        pd.DataFrame(pedidos_compra)
        if pedidos_compra
        else pd.DataFrame(
            columns=[
                "material", "descricao", "classe",
                "periodo_necessidade", "periodo_entrega", "quantidade",
            ]
        )
    )

    salvar(df_mrp, "03_mrp_projetado.csv")
    salvar(df_ped, "04_pedidos_compra.csv")

    # ── Sumário por material ──────────────────────────────────────────────────
    print(f"\n  {'Material':<10} {'Classe':>6} {'Pedidos Gerados':>15} {'Qtd Total':>12}")
    print(f"  {'─'*10} {'─'*6} {'─'*15} {'─'*12}")
    for mat in todos_mats:
        cls     = classe_map.get(mat, "C")
        ped_mat = df_ped[df_ped["material"] == mat] if not df_ped.empty else pd.DataFrame()
        n_ped   = len(ped_mat)
        qtd     = int(ped_mat["quantidade"].sum()) if not ped_mat.empty else 0
        print(f"  {mat:<10} {cls:>6} {n_ped:>15} {qtd:>12,}")

    return df_mrp, df_ped


# ─────────────────────────────────────────────────────────────────────────────
# PASSO 12: RATEIO FINAL POR ENDEREÇO
# ─────────────────────────────────────────────────────────────────────────────
def passo_12_rateio(df_pedidos: pd.DataFrame, estoque_detalhe: pd.DataFrame) -> pd.DataFrame:
    separador("PASSO 12 │ RATEIO FINAL POR ENDEREÇO")

    if df_pedidos.empty or df_pedidos["quantidade"].sum() == 0:
        print("  Nenhum pedido gerado — rateio não aplicável.")
        return pd.DataFrame()

    # Proporção de cada endereço dentro do material
    tot_mat = (
        estoque_detalhe.groupby("material", as_index=False)["quantidade"]
        .sum()
        .rename(columns={"quantidade": "total_mat"})
    )
    est_prop = estoque_detalhe.merge(tot_mat, on="material")

    def _proporcao(row: pd.Series) -> float:
        if row["total_mat"] > 0:
            return row["quantidade"] / row["total_mat"]
        # Material com estoque zero: dividir igualmente pelos endereços
        n = len(estoque_detalhe[estoque_detalhe["material"] == row["material"]])
        return 1.0 / max(n, 1)

    est_prop["proporcao"] = est_prop.apply(_proporcao, axis=1)

    # Agrupar pedidos por material + período de entrega
    ped_agr = df_pedidos.groupby(["material", "periodo_entrega"], as_index=False)["quantidade"].sum()

    linhas: list[dict] = []
    for _, ped in ped_agr.iterrows():
        mat        = ped["material"]
        qtd_total  = ped["quantidade"]
        per_ent    = ped["periodo_entrega"]

        end_mat = est_prop[est_prop["material"] == mat]

        if end_mat.empty:
            # Material sem endereço registrado → deposito geral
            linhas.append(
                {
                    "material"      : mat,
                    "periodo_entrega": per_ent,
                    "endereco"      : "DEPOSITO-GERAL",
                    "qtd_rateada"   : qtd_total,
                    "proporcao_pct" : 100.0,
                }
            )
        else:
            for _, e in end_mat.iterrows():
                linhas.append(
                    {
                        "material"      : mat,
                        "periodo_entrega": per_ent,
                        "endereco"      : e["endereco"],
                        "qtd_rateada"   : round(qtd_total * e["proporcao"], 2),
                        "proporcao_pct" : round(e["proporcao"] * 100, 2),
                    }
                )

    df_rateio = pd.DataFrame(linhas)
    salvar(df_rateio, "05_rateio_final.csv")

    print()
    print(df_rateio.to_string(index=False))
    return df_rateio


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def _imprimir_mrp_pivot(df_mrp: pd.DataFrame) -> None:
    """Exibe resumo do MRP projetado no console (estoque projetado por período)."""
    separador("RESUMO │ ESTOQUE PROJETADO POR MATERIAL E PERÍODO")
    pivot = df_mrp.pivot_table(
        index=["material", "classe"],
        columns="periodo",
        values="estoque_projetado",
        aggfunc="sum",
    )
    # Mostrar apenas os primeiros 6 meses para não sobrecarregar o console
    cols = list(pivot.columns)[:6]
    print(pivot[cols].to_string())
    if len(pivot.columns) > 6:
        print(f"  ... (+{len(pivot.columns) - 6} períodos adicionais no CSV)")


def main() -> None:
    cabecalho("SISTEMA MRP COM ENDEREÇAMENTO DE ESTOQUE")
    print(f"  Data de execução : {date.today().strftime('%d/%m/%Y')}")
    print(f"  Horizonte        : {HORIZONTE_MESES} meses")
    print(f"  Lead time        : {LEAD_TIME_MESES} mês(es)")
    print(f"  Estoque segurança: cobertura de {MESES_COBERTURA_SS} meses (rolling)")
    print(f"  Limite Classe C  : máximo {MAX_PEDIDOS_CLASSE_C} pedidos/ano")

    os.makedirs(DIR_DADOS, exist_ok=True)
    os.makedirs(DIR_SAIDA, exist_ok=True)

    # ── Carregar master de materiais ──────────────────────────────────────────
    materiais = pd.read_csv(os.path.join(DIR_DADOS, "materiais.csv"))

    # ── Executar passos na ordem obrigatória ──────────────────────────────────
    demanda               = passo_1_2_demanda()
    estoque, est_detalhe  = passo_3_estoque()
    entradas              = passo_4_pedidos_abertos()
    abc                   = passo_5_abc(demanda, materiais)
    df_mrp, df_ped        = passos_6_11_mrp(demanda, estoque, entradas, abc, materiais)
    rateio                = passo_12_rateio(df_ped, est_detalhe)

    _imprimir_mrp_pivot(df_mrp)

    # ── Resumo final ──────────────────────────────────────────────────────────
    cabecalho("RESUMO FINAL")
    print(f"  Materiais processados : {df_mrp['material'].nunique()}")
    print(f"  Períodos planejados   : {df_mrp['periodo'].nunique()}")
    print(f"  Pedidos de compra     : {len(df_ped)}")

    if not df_ped.empty:
        print(f"  Volume total (un)     : {df_ped['quantidade'].sum():,.0f}")
        print()
        print(f"  Por classe:")
        for cls in ["A", "B", "C"]:
            sub = df_ped[df_ped["classe"] == cls]
            if not sub.empty:
                print(
                    f"    Classe {cls} → {len(sub):>3} pedido(s) │ "
                    f"{sub['quantidade'].sum():>8,.0f} unidades"
                )

    print(f"\n  Arquivos gerados em ./{DIR_SAIDA}/")
    for arq in sorted(os.listdir(DIR_SAIDA)):
        if arq.endswith(".csv"):
            print(f"    ✓ {arq}")

    print(f"\n{'═' * W}\n")


if __name__ == "__main__":
    main()
