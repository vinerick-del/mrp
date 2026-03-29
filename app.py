"""
app.py — Interface Streamlit para o Sistema MRP SAP
====================================================
Upload dos 4 arquivos SAP + Lead Times → Processar → Dashboard + Download Excel

Persistência: arquivos importados são salvos em DIR_DADOS e recarregados
automaticamente na próxima abertura do app.
"""

import io
import os
import tempfile
from datetime import date, datetime, timedelta

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
# HELPERS DE PERSISTÊNCIA
# ─────────────────────────────────────────────────────────────────────────────
_ARQUIVOS_MAPA = {
    "demanda"   : ARQUIVO_DEMANDA_RAW,
    "remessas"  : "remessas_sap.csv",
    "pedidos"   : "pedidos_abertos.csv",
    "estoque"   : "estoque_sap.csv",
    "contratos" : "contratos_sap.csv",
    "materiais" : "materiais.csv",
    "lead"      : "lead_times.csv",
    "mb51"      : "historico_mb51.csv",
    "politica"  : "politica_pagamento.csv",
}


def _path_arquivo(chave: str) -> str:
    return os.path.join(DIR_DADOS, _ARQUIVOS_MAPA[chave])


def _info_arquivo(chave: str) -> str | None:
    """Retorna string com data de modificação ou None se não existir."""
    p = _path_arquivo(chave)
    if not os.path.exists(p):
        return None
    mtime = datetime.fromtimestamp(os.path.getmtime(p))
    return mtime.strftime("%d/%m/%Y %H:%M")


def _salvar_upload(uploaded, chave: str) -> bool:
    """Salva UploadedFile em DIR_DADOS. Retorna True se ok (ou se nada enviado)."""
    if not uploaded:
        return True
    path = _path_arquivo(chave)
    os.makedirs(DIR_DADOS, exist_ok=True)
    try:
        uploaded.seek(0)
        with open(path, "wb") as fh:
            fh.write(uploaded.read())
        return True
    except PermissionError:
        st.error(
            f"Sem permissão para salvar **{_ARQUIVOS_MAPA[chave]}**. "
            "Feche o arquivo no Excel e tente novamente."
        )
        return False


def _tem_dados_minimos() -> bool:
    """True se os arquivos obrigatórios (demanda + pelo menos 1 de pedidos) existem."""
    return os.path.exists(_path_arquivo("demanda")) and (
        os.path.exists(_path_arquivo("pedidos")) or os.path.exists(_path_arquivo("remessas"))
    )


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
# SIDEBAR — UPLOADS E STATUS DE PERSISTÊNCIA
# ─────────────────────────────────────────────────────────────────────────────
def _label_upload(numero: str, descricao: str, chave: str) -> str:
    info = _info_arquivo(chave)
    if info:
        return f"{numero} {descricao} ✅"
    return f"{numero} {descricao}"


def _caption_arquivo(chave: str):
    info = _info_arquivo(chave)
    if info:
        st.caption(f"Último import: {info}")


with st.sidebar:
    st.header("📂 Arquivos de Entrada")
    st.caption("Arquivos importados são **salvos automaticamente**. "
               "Na próxima abertura o app carrega os dados da última importação.")

    f_demanda   = st.file_uploader(_label_upload("①","Demanda (DTM)","demanda"),
                                   type=["csv","txt"], key="up_demanda",
                                   help="demanda_dtm_raw.csv — separado por ';'")
    _caption_arquivo("demanda")

    f_remessas  = st.file_uploader(_label_upload("②","Remessas SAP","remessas"),
                                   type=["csv","txt"], key="up_remessas",
                                   help="ME2M/ME9F — lookup de datas e nºs de documento")
    _caption_arquivo("remessas")

    f_pedidos   = st.file_uploader(_label_upload("③","Pedidos em Aberto","pedidos"),
                                   type=["csv","txt"], key="up_pedidos",
                                   help="pedidos_abertos.csv — base principal de qtd/valores")
    _caption_arquivo("pedidos")

    f_estoque   = st.file_uploader(_label_upload("④","Estoque SAP","estoque"),
                                   type=["csv","txt"], key="up_estoque",
                                   help="MB52/MMBE — separado por TAB")
    _caption_arquivo("estoque")

    f_contratos = st.file_uploader(_label_upload("⑤","Contratos SAP","contratos"),
                                   type=["csv","txt"], key="up_contratos",
                                   help="ME3M/ME3N — separado por TAB")
    _caption_arquivo("contratos")

    f_materiais = st.file_uploader(_label_upload("⑥","Materiais (catálogo)","materiais"),
                                   type=["csv","txt"], key="up_materiais",
                                   help="MM60/MM03 — CÓDIGO | DESCRIÇÃO | VALOR UNITÁRIO")
    _caption_arquivo("materiais")

    f_lead      = st.file_uploader(_label_upload("⑦","Lead Times","lead"),
                                   type=["csv"], key="up_lead",
                                   help="CSV: material,lead_time_dias")
    _caption_arquivo("lead")

    f_mb51      = st.file_uploader(_label_upload("⑧","Histórico MB51","mb51"),
                                   type=["csv","txt"], key="up_mb51",
                                   help="MB51 — movimentos 101/102")
    _caption_arquivo("mb51")

    f_politica  = st.file_uploader(_label_upload("⑨","Política de Pagamento","politica"),
                                   type=["csv","txt"], key="up_politica",
                                   help="CSV: documento | dias_parcela_1 | dias_parcela_2 ...")
    _caption_arquivo("politica")

    st.divider()
    btn_processar = st.button("🚀 Processar MRP", type="primary", use_container_width=True)
    if _tem_dados_minimos():
        st.caption("Dados disponíveis — processamento automático ativo.")

    st.divider()
    st.caption("**Fluxo:** Arquivos SAP → Normalização → Motor MRP → Dashboard + Excel")


# ─────────────────────────────────────────────────────────────────────────────
# SALVAR NOVOS UPLOADS IMEDIATAMENTE (antes do processamento)
# ─────────────────────────────────────────────────────────────────────────────
_uploads = {
    "demanda"  : f_demanda,
    "remessas" : f_remessas,
    "pedidos"  : f_pedidos,
    "estoque"  : f_estoque,
    "contratos": f_contratos,
    "materiais": f_materiais,
    "lead"     : f_lead,
    "mb51"     : f_mb51,
    "politica" : f_politica,
}
_novos_uploads = [k for k, v in _uploads.items() if v is not None]
if _novos_uploads:
    for chave, uploaded in _uploads.items():
        if uploaded:
            _salvar_upload(uploaded, chave)
    # Forçar reprocessamento quando novos arquivos chegarem
    st.session_state.pop("resultado", None)
    st.session_state.pop("_auto_processado", None)

# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO — disparado pelo botão OU automaticamente na primeira sessão
# ─────────────────────────────────────────────────────────────────────────────
_auto = _tem_dados_minimos() and "resultado" not in st.session_state and not st.session_state.get("_auto_processado")
_disparar = btn_processar or _auto

if _disparar:
    st.session_state["_auto_processado"] = True
    with st.spinner("Processando MRP..."):
        try:
            os.makedirs(DIR_DADOS, exist_ok=True)
            os.makedirs(DIR_SAIDA, exist_ok=True)

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

            # ── Cálculos Financeiros (independente de qual aba o usuário visita) ──
            _linhas_fin: list[pd.DataFrame] = []

            if not df_ped.empty:
                _tmp = df_ped.copy()
                print(f"  [FIN] Novos pedidos MRP      : R$ {_tmp['valor_total_pedido'].sum():,.2f}"
                      f" ({len(_tmp)} pedidos)")
                _tmp["mes_pedido"]           = pd.to_datetime(_tmp["data_pedido"], format="%d/%m/%Y", errors="coerce").dt.to_period("M").astype(str)
                _tmp["mes_entrega"]          = _tmp["periodo_entrega"]
                _tmp["data_base_pagamento"]  = pd.to_datetime(_tmp["data_chegada"], format="%d/%m/%Y", errors="coerce")
                _tmp["documento_referencia"] = None
                _tmp["numero_pedido"]        = None
                _tmp["origem"]               = "Novo Pedido (MRP)"
                _tmp["valor_pedido"]         = _tmp["valor_total_pedido"]
                _linhas_fin.append(_tmp[["origem","material","quantidade","valor_pedido",
                                         "mes_pedido","mes_entrega","data_base_pagamento",
                                         "documento_referencia","numero_pedido"]])

            if not df_abertos_fut.empty:
                _tmp2 = df_abertos_fut.copy()
                _val_sap = _tmp2["valor_total_pedido"].fillna(0).sum() if "valor_total_pedido" in _tmp2.columns else 0.0
                print(f"  [FIN] Pedidos existentes SAP : R$ {_val_sap:,.2f}")
                if "valor_total_pedido" not in _tmp2.columns or _val_sap == 0:
                    print("  [FIN] ⚠ Usando fallback ABC para precificar pedidos existentes")
                    _abc_price = abc[["material","valor_unitario"]].drop_duplicates("material")
                    _tmp2 = _tmp2.merge(_abc_price, on="material", how="left")
                    _tmp2["valor_total_pedido"] = _tmp2["quantidade"] * _tmp2["valor_unitario"].fillna(0)
                    print(f"  [FIN] Após fallback ABC       : R$ {_tmp2['valor_total_pedido'].sum():,.2f}")
                # mes_pedido = mês de emissão do PO (Data do documento) → visão orçamentária
                _tmp2["mes_pedido"] = (
                    _tmp2["mes_pedido"].fillna("Sem Data de Emissão")
                    if "mes_pedido" in _tmp2.columns else "Sem Data de Emissão"
                )
                _tmp2["mes_entrega"] = _tmp2["mes_remessa"]
                # data_base_pagamento = data de entrega real (data_remessa)
                if "data_remessa" in _tmp2.columns:
                    _tmp2["data_base_pagamento"] = pd.to_datetime(_tmp2["data_remessa"], errors="coerce")
                else:
                    _tmp2["data_base_pagamento"] = pd.to_datetime(
                        _tmp2["mes_remessa"] + "-01", format="%Y-%m-%d", errors="coerce"
                    )
                if "contrato" in _tmp2.columns:
                    _tmp2["documento_referencia"] = _tmp2["contrato"].astype(str).str.strip()
                elif "numero_pedido" in _tmp2.columns:
                    _tmp2["documento_referencia"] = _tmp2["numero_pedido"].astype(str).str.strip()
                else:
                    _tmp2["documento_referencia"] = None
                _tmp2["origem"]        = "Pedido Existente (SAP)"
                _tmp2["valor_pedido"]  = _tmp2["valor_total_pedido"].fillna(0)
                _tmp2["numero_pedido"] = _tmp2["numero_pedido"].astype(str).str.strip() if "numero_pedido" in _tmp2.columns else None
                _linhas_fin.append(_tmp2[["origem","material","quantidade","valor_pedido",
                                          "mes_pedido","mes_entrega","data_base_pagamento",
                                          "documento_referencia","numero_pedido"]])

            if not df_mb51.empty:
                _tmp3 = df_mb51.copy()
                _tmp3["mes_pedido"]          = _tmp3["mes_entrega"]
                _tmp3["data_base_pagamento"] = pd.to_datetime(_tmp3["mes_entrega"] + "-01", format="%Y-%m-%d", errors="coerce")
                _tmp3["documento_referencia"] = None
                _tmp3["numero_pedido"]        = None
                _tmp3["origem"]               = "Histórico Recebido (MB51)"
                _linhas_fin.append(_tmp3[["origem","material","quantidade","valor_pedido",
                                          "mes_pedido","mes_entrega","data_base_pagamento",
                                          "documento_referencia","numero_pedido"]])

            _df_fin = pd.concat(_linhas_fin, ignore_index=True) if _linhas_fin else pd.DataFrame()

            # Visão Orçamentária
            _vis_orc = pd.DataFrame()
            if not _df_fin.empty:
                _vis_orc = (
                    _df_fin.groupby(["mes_pedido","origem"], as_index=False)["valor_pedido"]
                    .sum()
                    .pivot(index="mes_pedido", columns="origem", values="valor_pedido")
                    .fillna(0).sort_index()
                )
                _vis_orc["Total"] = _vis_orc.sum(axis=1)

            # Visão de Caixa — explodir parcelas de pagamento
            _vis_cx = pd.DataFrame()
            _log_sem_politica: list[dict] = []
            if not _df_fin.empty:
                _parcelas: list[dict] = []
                for _, _row in _df_fin.iterrows():
                    _data_base = _row["data_base_pagamento"]
                    if pd.isna(_data_base):
                        continue
                    _contrato_ref = _row["documento_referencia"]
                    _pedido_ref   = _row.get("numero_pedido")
                    _dias = None
                    if _contrato_ref and str(_contrato_ref) in politica_pag_carregada:
                        _dias = politica_pag_carregada[str(_contrato_ref)]
                    if _dias is None and _pedido_ref and str(_pedido_ref) in politica_pag_carregada:
                        _dias = politica_pag_carregada[str(_pedido_ref)]
                    if _dias is None:
                        _dias = [60, 90]
                        _log_sem_politica.append({
                            "origem"      : _row["origem"],
                            "material"    : _row["material"],
                            "contrato"    : _contrato_ref,
                            "pedido"      : _pedido_ref,
                            "valor_pedido": _row["valor_pedido"],
                        })
                    _n          = len(_dias)
                    _vbase      = _row["valor_pedido"] // _n
                    _resto      = _row["valor_pedido"] - _vbase * _n
                    for _i, _d in enumerate(_dias):
                        _parcelas.append({
                            "origem"       : _row["origem"],
                            "material"     : _row["material"],
                            "mes_emissao"  : _row.get("mes_pedido", "-"),
                            "mes_entrega"  : _row.get("mes_entrega", "-"),
                            "prazo_dias"   : _d,
                            "mes_pagamento": (_data_base + timedelta(days=_d)).strftime("%Y-%m"),
                            "valor_parcela": _vbase + (_resto if _i == _n - 1 else 0),
                        })
                if _parcelas:
                    _df_fluxo = pd.DataFrame(_parcelas)
                    _vis_cx = (
                        _df_fluxo
                        .groupby(["mes_pagamento","origem"], as_index=False)["valor_parcela"]
                        .sum()
                        .pivot(index="mes_pagamento", columns="origem", values="valor_parcela")
                        .fillna(0).sort_index()
                    )
                    _vis_cx["Total"] = _vis_cx.sum(axis=1)

            # Renomear colunas do MRP para o formato solicitado
            df_mrp_out = df_mrp.rename(columns={
                "periodo"           : "mes",
                "total_entradas"    : "entrada",
                "estoque_projetado" : "estoque_proj",
                "necessidade"       : "necessidade",
            })[["material", "mes", "demanda", "entrada", "estoque_proj",
                "necessidade", "classe", "pedido_gerado"]]

            st.session_state["resultado"] = {
                "mrp"              : df_mrp_out,
                "pedidos"          : df_ped,
                "abertos_fut"      : df_abertos_fut,
                "historico_mb51"   : df_mb51,
                "politica_pag"     : politica_pag_carregada,
                "vis_orcamentaria" : _vis_orc,
                "vis_caixa"        : _vis_cx,
                "detalhe_fluxo"    : _df_fluxo if _parcelas else pd.DataFrame(),
                "log_sem_politica" : _log_sem_politica,
                "rateio"        : df_rateio,
                "alertas_rup"   : alertas_rup,
                "alertas_cont"  : alertas_cont,
                "contratos"     : contratos,
                "abc"           : abc,
                # dados auxiliares para Projeção de Estoque
                "estoque_inicial"  : estoque,
                "lead_times_dict"  : lt_dict or {},
                "materiais_df"     : materiais,
                "demanda_df"       : demanda,
            }
            if _auto and not btn_processar:
                st.info(
                    f"📂 Dados carregados automaticamente — "
                    f"{df_mrp['material'].nunique()} material(is) · "
                    f"{df_mrp['periodo'].nunique()} meses · "
                    f"{len(df_ped)} pedido(s)"
                )
            else:
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
    df_mrp         = r["mrp"]
    df_ped         = r["pedidos"]
    df_abertos_fut = r.get("abertos_fut", pd.DataFrame())
    df_mb51        = r.get("historico_mb51", pd.DataFrame())
    df_rateio      = r["rateio"]
    rup            = r["alertas_rup"]
    cont           = r["alertas_cont"]
    contratos      = r["contratos"]
    vis_orc        = r.get("vis_orcamentaria", pd.DataFrame())
    vis_cx         = r.get("vis_caixa",        pd.DataFrame())
    detalhe_fluxo  = r.get("detalhe_fluxo",   pd.DataFrame())
    log_sem_pol    = r.get("log_sem_politica", [])

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
        st.subheader("Projeção de Estoque")

        # ── Dados auxiliares ──────────────────────────────────────────────────
        _est_ini_df  = r.get("estoque_inicial", pd.DataFrame())
        _lt_dict     = r.get("lead_times_dict", {})
        _mat_df      = r.get("materiais_df", pd.DataFrame())
        _dem_df      = r.get("demanda_df", pd.DataFrame())
        _abc         = r["abc"]

        # Média de demanda mensal por material
        _dem_avg = (
            df_mrp.groupby("material")["demanda"].mean().rename("avg_dem").reset_index()
        )

        # Base por material: classe + avg_dem + SS + MAX + lead + desc + est_ini
        _meta = (
            _abc[["material", "classe"]].drop_duplicates()
            .merge(_dem_avg, on="material", how="left")
        )
        _meta["avg_dem"] = _meta["avg_dem"].fillna(0)
        _meta["EST.SEG."] = (_meta["avg_dem"] * MESES_COBERTURA_SS).round(0).astype(int)
        _meta["EST.MÁX."] = (_meta["avg_dem"] * (MESES_COBERTURA_SS + 1)).round(0).astype(int)
        _meta["LEAD(D)"]  = _meta["material"].apply(
            lambda m: _lt_dict.get(str(m), LEAD_TIME_DIAS)
        )

        if not _mat_df.empty and "descricao" in _mat_df.columns:
            _meta = _meta.merge(_mat_df[["material", "descricao"]], on="material", how="left")
        else:
            _meta["descricao"] = ""
        _meta["descricao"] = _meta["descricao"].fillna("")

        if not _est_ini_df.empty and "estoque_total" in _est_ini_df.columns:
            _meta = _meta.merge(
                _est_ini_df[["material", "estoque_total"]].rename(
                    columns={"estoque_total": "EST.INI"}
                ),
                on="material", how="left",
            )
            _meta["EST.INI"] = _meta["EST.INI"].fillna(0).round(0).astype(int)
        else:
            _meta["EST.INI"] = 0

        # Pivot estoque projetado por mês
        _proj = df_mrp.pivot_table(
            index="material", columns="mes", values="estoque_proj", aggfunc="sum"
        ).reset_index()

        # Renomear meses → "MMM/YYYY" em PT-BR
        _MESES_PT = {1:"JAN",2:"FEV",3:"MAR",4:"ABR",5:"MAI",6:"JUN",
                     7:"JUL",8:"AGO",9:"SET",10:"OUT",11:"NOV",12:"DEZ"}
        _mes_cols_orig = [c for c in _proj.columns if c != "material"]
        _mes_rename = {}
        for m in _mes_cols_orig:
            try:
                dt = pd.to_datetime(m + "-01")
                _mes_rename[m] = f"{_MESES_PT[dt.month]}/{dt.year}"
            except Exception:
                _mes_rename[m] = m
        _proj = _proj.rename(columns=_mes_rename)
        _mes_labels = [_mes_rename[m] for m in _mes_cols_orig]

        # Juntar tudo
        _display = _meta.merge(_proj, on="material", how="left")

        # ── Filtros ───────────────────────────────────────────────────────────
        _cf1, _cf2, _cf3, _cf4 = st.columns([2, 2, 3, 1])
        _classes_opts = ["Todas as Classes"] + sorted(_display["classe"].dropna().unique())
        _sel_cls  = _cf1.selectbox("Classe", _classes_opts, label_visibility="collapsed")

        _status_opts = ["Todos os Status", "Com ruptura", "Sem ruptura", "Abaixo do SS"]
        _sel_status = _cf2.selectbox("Status", _status_opts, label_visibility="collapsed")

        _search = _cf3.text_input("Código ou descrição...", "", label_visibility="collapsed",
                                   placeholder="Código ou descrição...")

        # Aplicar filtros
        _df_f = _display.copy()
        if _sel_cls != "Todas as Classes":
            _df_f = _df_f[_df_f["classe"] == _sel_cls]
        if _search:
            _mask = (
                _df_f["material"].astype(str).str.contains(_search, case=False, na=False) |
                _df_f["descricao"].astype(str).str.contains(_search, case=False, na=False)
            )
            _df_f = _df_f[_mask]
        if _sel_status == "Com ruptura":
            _rupt_mask = (_df_f[_mes_labels].fillna(0) <= 0).any(axis=1)
            _df_f = _df_f[_rupt_mask]
        elif _sel_status == "Sem ruptura":
            _ok_mask = (_df_f[_mes_labels].fillna(0) > 0).all(axis=1)
            _df_f = _df_f[_ok_mask]
        elif _sel_status == "Abaixo do SS":
            _ss_mask = (
                _df_f[_mes_labels].fillna(0)
                .lt(_df_f["EST.SEG."].values.reshape(-1, 1))
            ).any(axis=1)
            _df_f = _df_f[_ss_mask]

        _cf4.markdown(f"**{len(_df_f)}** materiais", unsafe_allow_html=True)

        # Banner informativo
        st.info(
            "As rupturas (células vermelhas) são **esperadas** — representam o período "
            "de trânsito do pedido (entre emissão e entrega). O Plano de Compras já contém "
            "os pedidos planejados para resolver essas rupturas.",
            icon="ℹ️",
        )

        # ── Tabela estilizada ─────────────────────────────────────────────────
        _cols_fixas = ["classe", "material", "descricao", "EST.INI", "EST.SEG.", "EST.MÁX.", "LEAD(D)"]
        _cols_view  = _cols_fixas + [m for m in _mes_labels if m in _df_f.columns]
        _tbl = _df_f[_cols_view].copy().reset_index(drop=True)

        # Truncar descrição longa
        _tbl["descricao"] = _tbl["descricao"].str[:40]

        # Renomear para exibição
        _tbl = _tbl.rename(columns={
            "classe"   : "CLS",
            "material" : "CÓDIGO",
            "descricao": "DESCRIÇÃO",
        })
        _mes_display = [m for m in _mes_labels if m in _cols_view]

        def _style_proj(row):
            styles = pd.Series("", index=row.index)
            ss = row.get("EST.SEG.", 0) or 0

            # Badge de classe
            _cls_bg = {"A": "background-color:#c0392b;color:white;font-weight:bold;text-align:center",
                       "B": "background-color:#e67e22;color:white;font-weight:bold;text-align:center",
                       "C": "background-color:#2980b9;color:white;font-weight:bold;text-align:center"}
            styles["CLS"] = _cls_bg.get(str(row.get("CLS", "")), "")

            # Colunas de referência
            styles["EST.SEG."] = "background-color:#b7e1b0;font-weight:bold"
            styles["EST.MÁX."] = "background-color:#d9f2d0"

            # Células mensais
            for col in _mes_display:
                if col not in row.index:
                    continue
                v = row[col]
                if pd.isna(v):
                    styles[col] = "color:#aaaaaa"
                elif v <= 0:
                    styles[col] = "background-color:#e74c3c;color:white;font-weight:bold"
                elif v < ss:
                    styles[col] = "background-color:#e67e22;color:white"
                elif v < ss * 1.15:
                    styles[col] = "background-color:#f9e04b"
                else:
                    styles[col] = "background-color:#b7e1b0"
            return styles

        _n_cells = _tbl.shape[0] * _tbl.shape[1]
        pd.set_option("styler.render.max_elements", max(_n_cells, 262144))

        _styled = (
            _tbl.style
            .apply(_style_proj, axis=1)
            .format(
                {c: lambda v: "—" if pd.isna(v) else f"{int(round(v)):,}".replace(",", ".")
                 for c in _mes_display},
                na_rep="—",
            )
            .format({"EST.INI": lambda v: f"{int(v):,}".replace(",","."),
                     "EST.SEG.": lambda v: f"{int(v):,}".replace(",","."),
                     "EST.MÁX.": lambda v: f"{int(v):,}".replace(",",".")},
                    na_rep="0")
        )

        st.dataframe(_styled, use_container_width=True, height=500)

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
        if vis_orc.empty and vis_cx.empty:
            st.info("Nenhum dado financeiro disponível.")
        else:
            subtab_orc, subtab_cx = st.tabs([
                "📊 Visão Orçamentária (Emissão)",
                "💸 Visão de Caixa (Desembolso Real)",
            ])

            with subtab_orc:
                st.caption("Compromisso financeiro agrupado por mês de emissão do pedido")
                if vis_orc.empty:
                    st.info("Sem dados orçamentários.")
                else:
                    _chart_financeiro(vis_orc, "Compromisso por Mês de Emissão")
                    agg_fmt = vis_orc.copy()
                    for col in agg_fmt.columns:
                        agg_fmt[col] = agg_fmt[col].apply(_fmt_brl_contabil)
                    st.dataframe(agg_fmt, use_container_width=True)

            with subtab_cx:
                st.caption(
                    "Desembolso previsto agrupado por mês de pagamento · "
                    "política padrão: 50% em 60 dias + 50% em 90 dias após recebimento"
                )
                if vis_cx.empty:
                    st.info("Nenhuma parcela de pagamento calculada.")
                else:
                    _chart_financeiro(vis_cx, "Desembolso por Mês de Pagamento")
                    agg_cx_fmt = vis_cx.copy()
                    for col in agg_cx_fmt.columns:
                        agg_cx_fmt[col] = agg_cx_fmt[col].apply(_fmt_brl_contabil)
                    st.dataframe(agg_cx_fmt, use_container_width=True)

                    # ── Rastreio: Emissão → Entrega → Pagamento ───────────────
                    with st.expander("🔍 Rastreio: Emissão → Entrega → Pagamento"):
                        st.caption(
                            "Cada linha representa uma parcela de pagamento. "
                            "Pagamentos em 2026/2027 podem ter origem em pedidos emitidos em anos anteriores "
                            "— rastreie pela coluna **Mês Emissão** (quando o PO foi criado) "
                            "e **Mês Entrega** (quando o material chega ao estoque)."
                        )
                        if not detalhe_fluxo.empty:
                            df_trace = (
                                detalhe_fluxo
                                .rename(columns={
                                    "origem"       : "Origem",
                                    "material"     : "Material",
                                    "mes_emissao"  : "Mês Emissão",
                                    "mes_entrega"  : "Mês Entrega",
                                    "prazo_dias"   : "Prazo (dias)",
                                    "mes_pagamento": "Mês Pagamento",
                                    "valor_parcela": "Valor Parcela",
                                })
                                .sort_values(["Mês Pagamento", "Mês Entrega"])
                                .reset_index(drop=True)
                            )
                            st.dataframe(
                                df_trace.style.format(
                                    {"Valor Parcela": _fmt_brl_contabil}, na_rep="-"
                                ),
                                use_container_width=True,
                                height=400,
                            )

                    with st.expander("⚠️ Log: Documentos sem Política (Aplicado Padrão 60/90 dias)"):
                        if log_sem_pol:
                            df_log = (
                                pd.DataFrame(log_sem_pol)
                                .drop_duplicates()
                                .reset_index(drop=True)
                            )
                            st.caption(f"{len(df_log)} documento(s) usaram a política padrão [60, 90] dias.")
                            st.dataframe(
                                df_log.style.format({"valor_pedido": _fmt_brl_contabil}, na_rep="-"),
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
    # Visão Financeira — sempre disponíveis (computadas no processamento)
    if not vis_orc.empty:
        dfs_excel["Visão Orçamentária"] = vis_orc
    if not vis_cx.empty:
        dfs_excel["Visão de Caixa"] = vis_cx
    if not detalhe_fluxo.empty:
        dfs_excel["Rastreio Pagamentos"] = detalhe_fluxo.sort_values(
            ["mes_pagamento", "mes_entrega"]
        ).reset_index(drop=True)

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
