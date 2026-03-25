"""
app.py — Interface Streamlit para o Sistema MRP SAP
====================================================
Upload dos 4 arquivos SAP + Lead Times → Processar → Dashboard + Download Excel
"""

import io
import os
import tempfile
from datetime import date

import pandas as pd
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
# HELPER: gerar Excel com múltiplas abas em memória
# ─────────────────────────────────────────────────────────────────────────────
def _gerar_excel(dfs: dict[str, pd.DataFrame]) -> bytes:
    """Recebe {nome_aba: DataFrame} e retorna bytes do .xlsx."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for nome_aba, df in dfs.items():
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
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
    df_mrp      = r["mrp"]
    df_ped      = r["pedidos"]
    df_rateio   = r["rateio"]
    rup         = r["alertas_rup"]
    cont        = r["alertas_cont"]
    contratos   = r["contratos"]

    # ── Métricas resumo ───────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Materiais", df_mrp["material"].nunique())
    c2.metric("Pedidos gerados", len(df_ped),
              delta=f"{int(df_ped['quantidade'].sum()):,} un." if not df_ped.empty else "0")
    c3.metric("⚠ Rupturas detectadas", len(rup["material"].unique()) if not rup.empty else 0)
    c4.metric("⚠ Contratos insuficientes", len(cont) if not cont.empty else 0)

    st.divider()

    # ── Abas do dashboard ─────────────────────────────────────────────────────
    tab_mrp, tab_proj, tab_ped, tab_rup, tab_cont, tab_rat = st.tabs([
        "📊 MRP Projetado",
        "📅 Projeção Mensal",
        "🛒 Pedidos a Gerar",
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
                col.metric(f"Classe {cls}", f"{len(sub)} pedido(s)",
                           f"{int(sub['quantidade'].sum()):,} un." if not sub.empty else "0 un.")

            st.dataframe(df_ped, use_container_width=True, height=380)

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
