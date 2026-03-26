"""
app.py — Interface Streamlit para o Sistema MRP SAP
====================================================
Upload dos 4 arquivos SAP + Lead Times → Processar → Dashboard + Download Excel
"""

import io
import os
import tempfile
from datetime import date, timedelta

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── Importar funções do motor MRP existente ────────────────────────────────────
from mrp import (
    LEAD_TIME_DIAS,
    HORIZONTE_MESES,
    MESES_COBERTURA_SS,
    CLASSE_C_COBERTURA_MESES,
    passo_1_2_demanda,
    passo_3_estoque,
    passo_4_pedidos_abertos,
    passo_5_abc,
    passos_6_11_mrp,
    passo_12_rateio,
    ler_remessas_sap,
    ler_estoque_sap,
    ler_contratos_sap,
    ler_lead_times,
    ler_materiais,
    ler_historico_mb51,
    ler_politica_pagamento,
    derivar_rateio_da_demanda,
    transformar_demanda_dtm,
    ARQUIVO_DEMANDA_RAW,
    DIR_DADOS,
    DIR_SAIDA,
)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sistema MRP",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📦 Sistema MRP — Planejamento de Materiais")
st.caption(
    f"Horizonte: **{HORIZONTE_MESES} meses** · "
    f"SS A/B: **{MESES_COBERTURA_SS}m rolling** · "
    f"Classe C cobertura: **{CLASSE_C_COBERTURA_MESES}m** · "
    f"Lead time default: **{LEAD_TIME_DIAS}d**"
)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: salvar uploaded file em temp e retornar path
# ─────────────────────────────────────────────────────────────────────────────
def _tmp_path(uploaded, suffix=".csv") -> str:
    """Salva UploadedFile em arquivo temporário e retorna o caminho."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.read())
    tmp.flush()
    return tmp.name


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS FINANCEIROS
# ─────────────────────────────────────────────────────────────────────────────
_FMT_MOEDA_EXCEL = '"R$ "#,##0.00'   # formato contábil R$ para openpyxl


def _fmt_brl_contabil(v) -> str:
    """Formata float em estilo contábil brasileiro: R$ 5.657.464,51 ou R$ (1.234,56)."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "-"
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ ({s})" if neg else f"R$ {s}"


def _chart_financeiro(agg: pd.DataFrame, titulo: str, cor_tendencia: str = "#FFD700") -> None:
    """Gráfico de barras empilhadas por origem + linha de tendência Total."""
    origens = [c for c in agg.columns if c != "Total"]
    fig = go.Figure()
    for orig in origens:
        fig.add_trace(go.Bar(name=orig, x=list(agg.index), y=list(agg[orig])))
    fig.add_trace(go.Scatter(
        name="Total (tendência)",
        x=list(agg.index),
        y=list(agg["Total"]),
        mode="lines+markers",
        line=dict(color=cor_tendencia, width=2, dash="dot"),
        marker=dict(size=7),
        yaxis="y",
    ))
    fig.update_layout(
        barmode="stack",
        title=titulo,
        yaxis_title="R$",
        legend=dict(orientation="h", y=-0.2),
        margin=dict(t=40, b=60),
        height=400,
    )
    st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: gerar Excel com múltiplas abas em memória
# ─────────────────────────────────────────────────────────────────────────────
def _gerar_excel(dfs: dict[str, pd.DataFrame]) -> bytes:
    """Recebe {nome_aba: DataFrame} e retorna bytes do .xlsx.
    Colunas cujo nome contém 'valor', 'total' ou 'parcela' recebem formato R$ contábil.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for nome_aba, df in dfs.items():
            df.to_excel(writer, sheet_name=nome_aba[:31], index=True)
            ws = writer.sheets[nome_aba[:31]]
            # Aplica formato contábil nas colunas de valor
            colunas_valor = [
                i + 2  # +1 porque index ocupa col A, +1 para base 1
                for i, col in enumerate(df.columns)
                if any(k in str(col).lower() for k in ("valor", "total", "parcela"))
            ]
            for col_idx in colunas_valor:
                for row in ws.iter_rows(
                    min_row=2, max_row=ws.max_row,
                    min_col=col_idx, max_col=col_idx
                ):
                    for cell in row:
                        cell.number_format = _FMT_MOEDA_EXCEL
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: calcular alertas a partir dos resultados MRP
# ─────────────────────────────────────────────────────────────────────────────
def _calcular_alertas(
    df_mrp: pd.DataFrame,
    df_ped: pd.DataFrame,
    contratos: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      alertas_ruptura   — material | periodo | estoque_projetado
      alertas_contrato  — material | pedido_total | saldo_contrato | deficit
    """
    # Ruptura: qualquer mês com estoque_projetado < 0
    alertas_rup = (
        df_mrp[df_mrp["estoque_projetado"] < 0][
            ["material", "periodo", "estoque_projetado"]
        ]
        .sort_values(["material", "periodo"])
        .reset_index(drop=True)
    )

    # Contrato insuficiente: total pedido pelo MRP > saldo_contrato
    alertas_cont = pd.DataFrame()
    if not df_ped.empty and not contratos.empty and "saldo_contrato" in contratos.columns:
        pedidos_total = (
            df_ped.groupby("material", as_index=False)["quantidade"]
            .sum()
            .rename(columns={"quantidade": "pedido_total"})
        )
        cmp = pedidos_total.merge(
            contratos[["material", "saldo_contrato"]],
            on="material",
            how="left",
        )
        cmp["saldo_contrato"] = cmp["saldo_contrato"].fillna(0)
        cmp["deficit"] = cmp["saldo_contrato"] - cmp["pedido_total"]
        alertas_cont = (
            cmp[cmp["deficit"] < 0][
                ["material", "pedido_total", "saldo_contrato", "deficit"]
            ]
            .sort_values("deficit")
            .reset_index(drop=True)
        )

    return alertas_rup, alertas_cont


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — UPLOADS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Arquivos de Entrada")

    f_demanda   = st.file_uploader("① Demanda (DTM format)", type=["csv", "txt"],
                                   help="Arquivo demanda_dtm_raw.csv — separado por ';'")
    f_remessas  = st.file_uploader("② Remessas SAP (entregas futuras)", type=["csv", "txt"],
                                   help="Exportação ME2M / ME9F — separado por TAB")
    f_estoque   = st.file_uploader("③ Estoque SAP (multi-depósito)", type=["csv", "txt"],
                                   help="Exportação MB52 / MMBE — separado por TAB")
    f_contratos = st.file_uploader("④ Contratos SAP (framework)", type=["csv", "txt"],
                                   help="Exportação ME3M / ME3N — separado por TAB")
    f_materiais = st.file_uploader("⑤ Materiais (catálogo SAP)", type=["csv", "txt"],
                                   help="Exportação MM60 / MM03 — CÓDIGO | DESCRIÇÃO | VALOR UNITÁRIO")
    f_lead      = st.file_uploader("⑥ Lead Times (opcional)", type=["csv"],
                                   help="CSV simples: material,lead_time_dias")
    f_mb51      = st.file_uploader("⑦ Histórico MB51 (Entradas 101/102)", type=["csv", "txt"],
                                   help="Relatório MB51 — movimentos 101 (recebimento) e 102 (estorno)")
    f_politica  = st.file_uploader("⑧ Política de Pagamento", type=["csv", "txt"],
                                   help="CSV: documento (contrato ou nº pedido) | dias_parcela_1 | dias_parcela_2 ...")

    st.divider()
    usar_dados_demo = st.checkbox("Usar dados existentes em data/", value=True,
                                  help="Processa os arquivos já presentes na pasta data/ do servidor")
    btn_processar = st.button("🚀 Processar MRP", type="primary", use_container_width=True)

    st.divider()
    st.caption(
        "**Fluxo:** Arquivos SAP → Normalização → "
        "Motor MRP existente → Dashboard + Excel"
    )


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────────────────────────────────────
if btn_processar:
    with st.spinner("Processando MRP..."):
        try:
            os.makedirs(DIR_DADOS, exist_ok=True)
            os.makedirs(DIR_SAIDA, exist_ok=True)

            # ── Salvar arquivos enviados pelo usuário (se houver) ─────────────
            def _salvar_upload(uploaded, nome_arquivo):
                """Salva o arquivo enviado em DIR_DADOS. Retorna True em sucesso."""
                if not uploaded:
                    return True
                path = os.path.join(DIR_DADOS, nome_arquivo)
                try:
                    with open(path, "wb") as fh:
                        fh.write(uploaded.read())
                    return True
                except PermissionError:
                    st.error(
                        f"Sem permissão para salvar **{nome_arquivo}**. "
                        "Feche o arquivo no Excel (ou outro programa) e tente novamente."
                    )
                    return False

            ok = all([
                _salvar_upload(f_demanda,   ARQUIVO_DEMANDA_RAW),
                _salvar_upload(f_remessas,  "remessas_sap.csv"),
                _salvar_upload(f_estoque,   "estoque_sap.csv"),
                _salvar_upload(f_contratos, "contratos_sap.csv"),
                _salvar_upload(f_materiais, "materiais.csv"),
                _salvar_upload(f_lead,      "lead_times.csv"),
                _salvar_upload(f_mb51,      "historico_mb51.csv"),
                _salvar_upload(f_politica,  "politica_pagamento.csv"),
            ])
            if not ok:
                st.stop()

            # ── Carregar materiais (legado — necessário para ABC fallback) ────
            mat_path = os.path.join(DIR_DADOS, "materiais.csv")
            materiais = ler_materiais(mat_path) if os.path.exists(mat_path) else pd.DataFrame(
                columns=["material", "descricao", "valor_unitario"]
            )

            # ── Contratos SAP ─────────────────────────────────────────────────
            cont_path = os.path.join(DIR_DADOS, "contratos_sap.csv")
            contratos = ler_contratos_sap(cont_path) if os.path.exists(cont_path) else pd.DataFrame()

            # ── Lead Times ────────────────────────────────────────────────────
            # Prioritário: coluna LEAD_TIME do materiais.csv
            # Fallback: lead_times.csv separado (upload ⑥)
            if "lead_time_dias" in materiais.columns:
                lt_dict = {
                    str(row["material"]): int(row["lead_time_dias"])
                    for _, row in materiais.iterrows()
                    if pd.notna(row["lead_time_dias"])
                }
            else:
                lt_path = os.path.join(DIR_DADOS, "lead_times.csv")
                lt_dict = ler_lead_times(lt_path) if os.path.exists(lt_path) else {}

            # ── Pipeline MRP ──────────────────────────────────────────────────
            demanda                  = passo_1_2_demanda()
            estoque                  = passo_3_estoque()
            entradas, df_abertos_fut = passo_4_pedidos_abertos()
            abc                      = passo_5_abc(demanda, materiais, contratos=contratos)
            df_mrp, df_ped           = passos_6_11_mrp(
                demanda, estoque, entradas, abc, materiais,
                lead_time_dias=LEAD_TIME_DIAS,
                lead_times_dict=lt_dict or None,
            )

            # Detalhe de demanda para rateio
            raw_path = os.path.join(DIR_DADOS, ARQUIVO_DEMANDA_RAW) if ARQUIVO_DEMANDA_RAW else None
            demanda_detail = (
                transformar_demanda_dtm(raw_path)
                if raw_path and os.path.exists(raw_path)
                else None
            )
            df_rateio = passo_12_rateio(df_ped, df_abertos_fut, demanda_detail=demanda_detail)

            # ── Histórico MB51 (opcional) ─────────────────────────────────────
            mb51_path = os.path.join(DIR_DADOS, "historico_mb51.csv")
            df_mb51 = ler_historico_mb51(mb51_path) if os.path.exists(mb51_path) else pd.DataFrame()

            # ── Política de Pagamento (opcional) ─────────────────────────────
            politica_path = os.path.join(DIR_DADOS, "politica_pagamento.csv")
            politica_pag_carregada: dict[str, list[int]] = (
                ler_politica_pagamento(politica_path)
                if os.path.exists(politica_path) else {}
            )

            # Alertas
            alertas_rup, alertas_cont = _calcular_alertas(df_mrp, df_ped, contratos)

            # Renomear colunas do MRP para o formato solicitado
            df_mrp_out = df_mrp.rename(columns={
                "periodo"           : "mes",
                "total_entradas"    : "entrada",
                "estoque_projetado" : "estoque_proj",
                "necessidade"       : "necessidade",
            })[["material", "mes", "demanda", "entrada", "estoque_proj",
                "necessidade", "classe", "pedido_gerado"]]

            st.session_state["resultado"] = {
                "mrp"           : df_mrp_out,
                "pedidos"       : df_ped,
                "abertos_fut"   : df_abertos_fut,
                "historico_mb51": df_mb51,
                "politica_pag"  : politica_pag_carregada,
                "rateio"        : df_rateio,
                "alertas_rup"   : alertas_rup,
                "alertas_cont"  : alertas_cont,
                "contratos"     : contratos,
                "abc"           : abc,
            }
            st.success(
                f"✅ MRP processado — {df_mrp['material'].nunique()} material(is) · "
                f"{df_mrp['periodo'].nunique()} meses · "
                f"{len(df_ped)} pedido(s) gerado(s)"
            )

        except Exception as exc:
            st.error(f"❌ Erro no processamento: {exc}")
            st.exception(exc)


# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD — exibido após processamento
# ─────────────────────────────────────────────────────────────────────────────
if "resultado" in st.session_state:
    r = st.session_state["resultado"]
    df_mrp        = r["mrp"]
    df_ped        = r["pedidos"]
    df_abertos_fut= r.get("abertos_fut", pd.DataFrame())
    df_mb51       = r.get("historico_mb51", pd.DataFrame())
    df_rateio     = r["rateio"]
    rup           = r["alertas_rup"]
    cont          = r["alertas_cont"]
    contratos     = r["contratos"]

    # ── Métricas resumo ───────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Materiais", df_mrp["material"].nunique())
    c2.metric("Pedidos gerados", len(df_ped),
              delta=f"{int(df_ped['quantidade'].sum()):,} un." if not df_ped.empty else "0")
    vol_total = df_ped["valor_total_pedido"].sum() if not df_ped.empty else 0.0
    if vol_total >= 1_000_000:
        vol_str = f"R$ {vol_total/1_000_000:.1f}M"
    elif vol_total >= 1_000:
        vol_str = f"R$ {vol_total/1_000:.1f}K"
    else:
        vol_str = f"R$ {vol_total:,.2f}"
    c3.metric("Volume financeiro", vol_str)
    c4.metric("⚠ Rupturas detectadas", len(rup["material"].unique()) if not rup.empty else 0)
    c5.metric("⚠ Contratos insuficientes", len(cont) if not cont.empty else 0)

    st.divider()

    # ── Abas do dashboard ─────────────────────────────────────────────────────
    tab_mrp, tab_proj, tab_ped, tab_fin, tab_rup, tab_cont, tab_rat = st.tabs([
        "📊 MRP Projetado",
        "📅 Projeção Mensal",
        "🛒 Pedidos a Gerar",
        "💰 Visão Financeira",
        "🔴 Alertas de Ruptura",
        "📋 Saldo de Contrato",
        "📂 Rateio",
    ])

    with tab_mrp:
        st.subheader("MRP Projetado")
        st.caption("material | mes | demanda | entrada | estoque_proj | necessidade | valor_pedido")

        # Enriquecer com valor_unitario do ABC para calcular valor do pedido gerado
        abc = r["abc"]
        df_mrp_val = df_mrp.merge(
            abc[["material", "valor_unitario"]].drop_duplicates("material"),
            on="material", how="left",
        )
        df_mrp_val["valor_unitario"] = df_mrp_val["valor_unitario"].fillna(0)
        df_mrp_val["valor_pedido"] = df_mrp_val["pedido_gerado"] * df_mrp_val["valor_unitario"]
        df_mrp_val = df_mrp_val.drop(columns=["valor_unitario"])

        mats = sorted(df_mrp_val["material"].unique())
        sel  = st.multiselect("Filtrar material(is)", mats, default=mats[:5] if len(mats) > 5 else mats)
        df_show = df_mrp_val[df_mrp_val["material"].isin(sel)] if sel else df_mrp_val

        n_cells = df_show.shape[0] * df_show.shape[1]
        pd.set_option("styler.render.max_elements", max(n_cells, 262144))
        st.dataframe(
            df_show.style.applymap(
                lambda v: "background-color: #ffcccc" if isinstance(v, (int, float)) and v < 0 else "",
                subset=["estoque_proj"],
            ),
            use_container_width=True,
            height=420,
        )

    # ── Helper: pivot projeção mensal (reutilizado na tela e no Excel) ──────────
    def _build_projecao_pivot(df: pd.DataFrame) -> pd.DataFrame:
        piv = (
            df.pivot_table(
                index=["material", "classe"],
                columns="mes",
                values="estoque_proj",
                aggfunc="sum",
            )
            .reset_index()
        )
        # Converter colunas "YYYY-MM" para datetime do 1º dia do mês
        mes_str = sorted([c for c in piv.columns if c not in ("material", "classe")])
        rename_map = {m: pd.to_datetime(m + "-01") for m in mes_str}
        piv = piv.rename(columns=rename_map)
        date_cols = [rename_map[m] for m in mes_str]
        piv = piv[["material", "classe"] + date_cols]
        if date_cols:
            piv["Saldo Final"] = piv[date_cols[-1]]
        return piv

    with tab_proj:
        st.subheader("Projeção Mensal de Estoque por Material")
        st.caption("Estoque projetado (un.) ao final de cada mês · última coluna = saldo final do horizonte")

        pivot = _build_projecao_pivot(df_mrp)
        date_cols = [c for c in pivot.columns if c not in ("material", "classe", "Saldo Final")]

        n_cells_piv = pivot.shape[0] * pivot.shape[1]
        pd.set_option("styler.render.max_elements", max(n_cells_piv, 262144))
        st.dataframe(
            pivot.style.applymap(
                lambda v: "background-color: #ffcccc" if isinstance(v, (int, float)) and v < 0 else "",
                subset=date_cols + (["Saldo Final"] if date_cols else []),
            ),
            use_container_width=True,
            height=450,
        )

    with tab_ped:
        st.subheader("Pedidos a Gerar")
        if df_ped.empty:
            st.info("Nenhum pedido gerado para o horizonte configurado.")
        else:
            c1p, c2p, c3p = st.columns(3)
            for cls, col in zip(["A", "B", "C"], [c1p, c2p, c3p]):
                sub = df_ped[df_ped["classe"] == cls]
                val = sub["valor_total_pedido"].sum() if not sub.empty else 0.0
                if val >= 1_000_000:
                    val_str = f"R$ {val/1_000_000:.1f}M"
                elif val >= 1_000:
                    val_str = f"R$ {val/1_000:.1f}K"
                else:
                    val_str = f"R$ {val:,.2f}"
                col.metric(
                    f"Classe {cls}",
                    f"{len(sub)} pedido(s)",
                    delta=f"{int(sub['quantidade'].sum()):,} un. · {val_str}" if not sub.empty else "0 un.",
                )

            fmt_moeda = {"valor_unitario": _fmt_brl_contabil, "valor_total_pedido": _fmt_brl_contabil}
            st.dataframe(
                df_ped.style.format(fmt_moeda, na_rep="-"),
                use_container_width=True,
                height=380,
            )

    with tab_fin:
        st.subheader("Visão Financeira")

        # ── Novos pedidos MRP ─────────────────────────────────────────────────
        # data_base = data_chegada calculada pelo lead time
        linhas_mrp = []
        if not df_ped.empty:
            tmp = df_ped.copy()
            tmp["mes_pedido"]           = pd.to_datetime(tmp["data_pedido"], format="%d/%m/%Y", errors="coerce").dt.to_period("M").astype(str)
            tmp["mes_entrega"]          = tmp["periodo_entrega"]
            tmp["data_base_pagamento"]  = pd.to_datetime(tmp["data_chegada"], format="%d/%m/%Y", errors="coerce")
            tmp["documento_referencia"] = None
            tmp["numero_pedido"]        = None
            tmp["origem"]               = "Novo Pedido (MRP)"
            tmp["valor_pedido"]         = tmp["valor_total_pedido"]
            linhas_mrp = [tmp[["origem", "material", "quantidade", "valor_pedido",
                                "mes_pedido", "mes_entrega",
                                "data_base_pagamento", "documento_referencia", "numero_pedido"]]]

        # ── Pedidos existentes SAP ────────────────────────────────────────────
        # data_base = data de remessa (previsão de entrega)
        linhas_sap = []
        if not df_abertos_fut.empty:
            tmp2 = df_abertos_fut.copy()
            if "valor_total_pedido" not in tmp2.columns or tmp2["valor_total_pedido"].fillna(0).sum() == 0:
                abc_price = r["abc"][["material", "valor_unitario"]].drop_duplicates("material")
                tmp2 = tmp2.merge(abc_price, on="material", how="left")
                tmp2["valor_total_pedido"] = tmp2["quantidade"] * tmp2["valor_unitario"].fillna(0)
            tmp2["mes_pedido"]  = tmp2.get("mes_pedido", pd.Series(["Já Comprometido"] * len(tmp2), index=tmp2.index))
            tmp2["mes_pedido"]  = tmp2["mes_pedido"].fillna("Já Comprometido")
            tmp2["mes_entrega"] = tmp2["mes_remessa"]
            tmp2["data_base_pagamento"]  = pd.to_datetime(
                tmp2["mes_remessa"] + "-01", format="%Y-%m-%d", errors="coerce"
            )
            # Prioridade: contrato (mais estável para política) > nº pedido
            if "contrato" in tmp2.columns:
                tmp2["documento_referencia"] = tmp2["contrato"].astype(str).str.strip()
            elif "numero_pedido" in tmp2.columns:
                tmp2["documento_referencia"] = tmp2["numero_pedido"].astype(str).str.strip()
            else:
                tmp2["documento_referencia"] = None
            tmp2["origem"]        = "Pedido Existente (SAP)"
            tmp2["valor_pedido"]  = tmp2["valor_total_pedido"].fillna(0)
            tmp2["numero_pedido"] = tmp2["numero_pedido"].astype(str).str.strip() if "numero_pedido" in tmp2.columns else None
            linhas_sap = [tmp2[["origem", "material", "quantidade", "valor_pedido",
                                 "mes_pedido", "mes_entrega",
                                 "data_base_pagamento", "documento_referencia", "numero_pedido"]]]

        # ── Histórico realizado MB51 ──────────────────────────────────────────
        # data_base = data de lançamento (1º dia do mês, pois já agregamos)
        linhas_mb51 = []
        if not df_mb51.empty:
            tmp3 = df_mb51.copy()
            # mes_pedido = mes_entrega: para histórico já recebido,
            # o mês de recebimento é o mês do compromisso orçamentário
            tmp3["mes_pedido"]          = tmp3["mes_entrega"]
            tmp3["data_base_pagamento"] = pd.to_datetime(
                tmp3["mes_entrega"] + "-01", format="%Y-%m-%d", errors="coerce"
            )
            tmp3["documento_referencia"] = None
            tmp3["numero_pedido"]        = None
            tmp3["origem"]               = "Histórico Recebido (MB51)"
            linhas_mb51 = [tmp3[["origem", "material", "quantidade", "valor_pedido",
                                  "mes_pedido", "mes_entrega",
                                  "data_base_pagamento", "documento_referencia", "numero_pedido"]]]

        todas = linhas_mrp + linhas_sap + linhas_mb51
        if not todas:
            st.info("Nenhum dado financeiro disponível.")
        else:
            df_fin = pd.concat(todas, ignore_index=True)

            # Política de pagamento: {documento_referencia: [dias]}
            politica_pag: dict[str, list[int]] = r.get("politica_pag", {})

            subtab_orc, subtab_cx = st.tabs([
                "📊 Visão Orçamentária (Emissão)",
                "💸 Visão de Caixa (Desembolso Real)",
            ])

            with subtab_orc:
                st.caption("Compromisso financeiro agrupado por mês de emissão do pedido")
                agg = (df_fin.groupby(["mes_pedido", "origem"], as_index=False)["valor_pedido"]
                       .sum()
                       .pivot(index="mes_pedido", columns="origem", values="valor_pedido")
                       .fillna(0)
                       .sort_index())
                agg["Total"] = agg.sum(axis=1)
                _chart_financeiro(agg, "Compromisso por Mês de Emissão")
                agg_fmt = agg.copy()
                for col in agg_fmt.columns:
                    agg_fmt[col] = agg_fmt[col].apply(_fmt_brl_contabil)
                st.dataframe(agg_fmt, use_container_width=True)
                # Guarda para exportação Excel
                st.session_state["_vis_orcamentaria"] = agg

            with subtab_cx:
                st.caption(
                    "Desembolso previsto agrupado por mês de pagamento · "
                    "política padrão: 50% em 60 dias + 50% em 90 dias após recebimento"
                )

                # ── Explodir parcelas de pagamento ────────────────────────────
                log_sem_politica: list[dict] = []
                parcelas: list[dict] = []

                for _, row in df_fin.iterrows():
                    data_base = row["data_base_pagamento"]
                    if pd.isna(data_base):
                        continue  # sem data de referência, não é possível calcular

                    contrato_ref = row["documento_referencia"]   # contrato SAP
                    pedido_ref   = row.get("numero_pedido")      # nº pedido SAP
                    dias_parcelas = None

                    # 1) busca pelo contrato
                    if contrato_ref and str(contrato_ref) in politica_pag:
                        dias_parcelas = politica_pag[str(contrato_ref)]

                    # 2) se não achou, busca pelo nº pedido
                    if dias_parcelas is None and pedido_ref and str(pedido_ref) in politica_pag:
                        dias_parcelas = politica_pag[str(pedido_ref)]

                    # 3) fallback: política padrão 60/90 dias
                    if dias_parcelas is None:
                        dias_parcelas = [60, 90]
                        log_sem_politica.append({
                            "origem"    : row["origem"],
                            "material"  : row["material"],
                            "contrato"  : contrato_ref,
                            "pedido"    : pedido_ref,
                            "valor_pedido": row["valor_pedido"],
                        })

                    # 3) gerar parcelas — soma bate exatamente com valor_pedido
                    n = len(dias_parcelas)
                    valor_base   = row["valor_pedido"] // n  # parte inteira
                    resto        = row["valor_pedido"] -  valor_base * n  # centavos restantes

                    for idx, d in enumerate(dias_parcelas):
                        data_pag    = data_base + timedelta(days=d)
                        mes_pag     = data_pag.strftime("%Y-%m")
                        valor_parc  = valor_base + (resto if idx == n - 1 else 0)
                        parcelas.append({
                            "origem"       : row["origem"],
                            "material"     : row["material"],
                            "mes_pagamento": mes_pag,
                            "valor_parcela": valor_parc,
                        })

                if not parcelas:
                    st.info("Nenhuma parcela de pagamento calculada.")
                else:
                    df_fluxo = pd.DataFrame(parcelas)

                    agg_cx = (
                        df_fluxo
                        .groupby(["mes_pagamento", "origem"], as_index=False)["valor_parcela"]
                        .sum()
                        .pivot(index="mes_pagamento", columns="origem", values="valor_parcela")
                        .fillna(0)
                        .sort_index()
                    )
                    agg_cx["Total"] = agg_cx.sum(axis=1)
                    _chart_financeiro(agg_cx, "Desembolso por Mês de Pagamento")
                    agg_cx_fmt = agg_cx.copy()
                    for col in agg_cx_fmt.columns:
                        agg_cx_fmt[col] = agg_cx_fmt[col].apply(_fmt_brl_contabil)
                    st.dataframe(agg_cx_fmt, use_container_width=True)
                    # Guarda para exportação Excel
                    st.session_state["_vis_caixa"] = agg_cx

                    with st.expander("⚠️ Log: Documentos sem Política (Aplicado Padrão 60/90 dias)"):
                        if log_sem_politica:
                            df_log = (
                                pd.DataFrame(log_sem_politica)
                                .drop_duplicates()
                                .reset_index(drop=True)
                            )
                            st.caption(f"{len(df_log)} documento(s) usaram a política padrão [60, 90] dias.")
                            st.dataframe(
                                df_log.style.format(
                                    {"valor_pedido": _fmt_brl_contabil}, na_rep="-"
                                ),
                                use_container_width=True,
                            )
                        else:
                            st.success("Todos os documentos possuem política de pagamento cadastrada.")

    with tab_rup:
        st.subheader("Alertas de Ruptura de Estoque")
        if rup.empty:
            st.success("✅ Nenhuma ruptura detectada no horizonte de planejamento.")
        else:
            st.error(f"⚠ {rup['material'].nunique()} material(is) com estoque projetado negativo")
            st.dataframe(rup, use_container_width=True)

    with tab_cont:
        st.subheader("Alertas de Saldo de Contrato Insuficiente")
        if contratos.empty:
            st.info("Arquivo de contratos SAP não carregado.")
        elif cont.empty:
            st.success("✅ Todos os pedidos gerados estão cobertos pelos contratos vigentes.")
        else:
            st.warning(f"⚠ {len(cont)} material(is) com saldo de contrato insuficiente")
            st.dataframe(
                cont.style.applymap(
                    lambda v: "background-color: #ffcccc" if isinstance(v, (int, float)) and v < 0 else "",
                    subset=["deficit"],
                ),
                use_container_width=True,
            )
            if not contratos.empty:
                st.caption("Contratos vigentes carregados:")
                st.dataframe(contratos, use_container_width=True)

    with tab_rat:
        st.subheader("Rateio por Departamento / Programa Orçamentário")
        if df_rateio.empty:
            st.info("Nenhum dado de rateio disponível.")
        else:
            st.dataframe(df_rateio, use_container_width=True, height=380)

    # ── Download Excel ────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📥 Download")

    dfs_excel: dict[str, pd.DataFrame] = {
        "MRP Projetado"   : df_mrp_val,
        "Projeção Mensal" : _build_projecao_pivot(df_mrp),
    }
    if not df_ped.empty:
        dfs_excel["Pedidos"] = df_ped
    if not df_rateio.empty:
        dfs_excel["Rateio"] = df_rateio
    if not rup.empty:
        dfs_excel["Alertas Ruptura"] = rup
    if not cont.empty:
        dfs_excel["Contrato Insuficiente"] = cont
    # Visão Financeira — adicionadas quando a aba é visitada
    if "_vis_orcamentaria" in st.session_state:
        dfs_excel["Visão Orçamentária"] = st.session_state["_vis_orcamentaria"]
    if "_vis_caixa" in st.session_state:
        dfs_excel["Visão de Caixa"] = st.session_state["_vis_caixa"]

    try:
        excel_bytes = _gerar_excel(dfs_excel)
        nome_arquivo = f"mrp_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇ Baixar Excel (todas as abas)",
            data=excel_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except ImportError:
        st.warning("openpyxl não instalado — execute: pip install openpyxl")
    except Exception as exc:
        st.error(f"Erro ao gerar Excel: {exc}")

else:
    # Estado inicial — instrução ao usuário
    st.info(
        "👈 Faça upload dos arquivos no painel lateral **ou** marque "
        "**'Usar dados existentes em data/'** e clique em **Processar MRP**."
    )
    with st.expander("ℹ️ Formato esperado de cada arquivo"):
        st.markdown("""
| Arquivo | Separador | Colunas-chave |
|---|---|---|
| ① Demanda DTM | `;` (ponto-e-vírgula) | CÓDIGO, MÊS, DEP., PROJETO, QTD |
| ② Remessas SAP | `TAB` | Material, Data de remessa, a ser fornecida (quantidade), Código de eliminação |
| ③ Estoque SAP | `TAB` | Produto, Qtd.disponível UMB |
| ④ Contratos SAP | `TAB` | Material, Fim da validade, Qtd.prev.pendente, Preço líquido |
| ⑤ Lead Times | `,` (vírgula) | material, lead_time_dias |

**Números em formato brasileiro:** `3.515,50` = 3515.50
**Datas:** `dd/mm/yyyy`
        """)
