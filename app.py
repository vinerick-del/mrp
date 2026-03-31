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
            # Prioridade: 1) coluna LT em materiais.csv  2) LEAD_TIMES.csv  3) lead_times.csv
            if "lead_time_dias" in materiais.columns:
                lt_dict = {
                    str(row["material"]): int(row["lead_time_dias"])
                    for _, row in materiais.iterrows()
                    if pd.notna(row["lead_time_dias"]) and row["lead_time_dias"] > 0
                }
                print(f"  [LT] Lead times de materiais.csv: {len(lt_dict)} itens")
            else:
                # Aceita qualquer capitalização do nome do arquivo
                _lt_candidatos = ["LEAD_TIMES.csv", "lead_times.csv", "Lead_Times.csv"]
                lt_path = next(
                    (os.path.join(DIR_DADOS, f) for f in _lt_candidatos
                     if os.path.exists(os.path.join(DIR_DADOS, f))),
                    None,
                )
                lt_dict = ler_lead_times(lt_path) if lt_path else {}
                if not lt_dict:
                    print(f"  [LT] ⚠ Nenhum lead time carregado — usando default {LEAD_TIME_DIAS}d para todos")
                    print(f"  [LT]   Colunas em materiais.csv: {list(materiais.columns)}")

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
            # demanda_detail salvo para filtros de departamento/programa na Projeção de Estoque
            _demanda_detail_df = demanda_detail if demanda_detail is not None else pd.DataFrame()

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
            _vis_cx   = pd.DataFrame()
            _df_fluxo = pd.DataFrame()
            _parcelas: list[dict] = []
            _log_sem_politica: list[dict] = []
            if not _df_fin.empty:
                _parcelas = []
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
                    _n     = len(_dias)
                    _vbase = round(_row["valor_pedido"] / _n, 2)
                    _resto = round(_row["valor_pedido"] - (_vbase * (_n - 1)), 2)
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

            # ── Rateio Financeiro: aplicar proporções dept/programa ────────────────
            # Base de proporções derivada do df_rateio (já calcula proporcao 0-1)
            _base_rateio = pd.DataFrame(
                columns=["material", "departamento", "programa_orcamentario", "proporcao"]
            )
            if not df_rateio.empty and "proporcao_pct" in df_rateio.columns:
                _base_rateio = (
                    df_rateio[["material", "departamento", "programa_orcamentario", "proporcao_pct"]]
                    .drop_duplicates()
                    .copy()
                )
                _base_rateio["proporcao"] = _base_rateio["proporcao_pct"] / 100
                _base_rateio = _base_rateio[
                    ["material", "departamento", "programa_orcamentario", "proporcao"]
                ]

            def _aplicar_rateio_fin(df: pd.DataFrame, col_valor: str) -> pd.DataFrame:
                if df.empty:
                    return pd.DataFrame()
                if _base_rateio.empty:
                    out = df.copy()
                    out["departamento"]          = "NÃO DEFINIDO"
                    out["programa_orcamentario"] = "NÃO DEFINIDO"
                    out["proporcao"]             = 1.0
                    out["valor_rateado"]         = out[col_valor]
                    return out
                merged = df.merge(_base_rateio, on="material", how="left")
                merged["departamento"]          = merged["departamento"].fillna("NÃO DEFINIDO")
                merged["programa_orcamentario"] = merged["programa_orcamentario"].fillna("NÃO DEFINIDO")
                merged["proporcao"]             = merged["proporcao"].fillna(1.0)
                merged["valor_rateado"]         = merged[col_valor] * merged["proporcao"]
                return merged

            _df_fin_bruto   = _aplicar_rateio_fin(_df_fin, "valor_pedido")
            _df_fluxo_bruto = _aplicar_rateio_fin(
                _df_fluxo if _parcelas else pd.DataFrame(), "valor_parcela"
            )

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
                "rateio"         : df_rateio,
                "df_fin_bruto"   : _df_fin_bruto,
                "df_fluxo_bruto" : _df_fluxo_bruto,
                "alertas_rup"    : alertas_rup,
                "alertas_cont"  : alertas_cont,
                "contratos"     : contratos,
                "abc"           : abc,
                # dados auxiliares para Projeção de Estoque
                "estoque_inicial"  : estoque,
                "lead_times_dict"  : lt_dict or {},
                "materiais_df"     : materiais,
                "demanda_df"       : demanda,
                "mrp_full"         : df_mrp,         # com estoque_seguranca_3m, entrada_pedidos_existentes
                "demanda_detail_df": _demanda_detail_df,  # com departamento e programa_orcamentario
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
    df_fin_bruto   = r.get("df_fin_bruto",    pd.DataFrame())
    df_fluxo_bruto = r.get("df_fluxo_bruto",  pd.DataFrame())

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
        "📅 Projeção de Estoque",
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
        import collections as _col_mod
        import streamlit.components.v1 as _stc

        # ── Dados auxiliares ──────────────────────────────────────────────────
        _est_ini_df   = r.get("estoque_inicial", pd.DataFrame())
        _lt_dict      = r.get("lead_times_dict", {})
        _mat_df       = r.get("materiais_df", pd.DataFrame())
        _abc          = r["abc"]
        _mrp_full     = r.get("mrp_full", pd.DataFrame())   # com estoque_seguranca_3m
        _dem_det      = r.get("demanda_detail_df", pd.DataFrame())  # material|mes|depto|prog
        _abertos      = r.get("abertos_fut", pd.DataFrame())
        _novos_ped    = r["pedidos"]

        _MESES_PT = {1:"JAN",2:"FEV",3:"MAR",4:"ABR",5:"MAI",6:"JUN",
                     7:"JUL",8:"AGO",9:"SET",10:"OUT",11:"NOV",12:"DEZ"}

        def _mes_label(m: str) -> str:
            try:
                dt = pd.to_datetime(m + "-01")
                return f"{_MESES_PT[dt.month]}/{dt.year}"
            except Exception:
                return m

        # ── Estoque de segurança por material: usa mrp_full se disponível ────
        if not _mrp_full.empty and "estoque_seguranca_3m" in _mrp_full.columns:
            _ss_map = (
                _mrp_full.groupby("material")["estoque_seguranca_3m"].mean()
                .round(0).astype(int).to_dict()
            )
        else:
            _dem_avg_map = df_mrp.groupby("material")["demanda"].mean().to_dict()
            _ss_map = {m: int(round(v * MESES_COBERTURA_SS)) for m, v in _dem_avg_map.items()}

        # ── Metadados por material ────────────────────────────────────────────
        _meta = _abc[["material", "classe"]].drop_duplicates().copy()
        _meta["EST.SEG."] = _meta["material"].map(_ss_map).fillna(0).astype(int)
        _meta["EST.MÁX."] = (_meta["EST.SEG."] * ((MESES_COBERTURA_SS + 1) / MESES_COBERTURA_SS)).round(0).astype(int)
        _meta["LEAD(D)"]  = _meta["material"].apply(lambda m: _lt_dict.get(str(m), LEAD_TIME_DIAS))

        if not _mat_df.empty and "descricao" in _mat_df.columns:
            _meta = _meta.merge(_mat_df[["material", "descricao"]], on="material", how="left")
        else:
            _meta["descricao"] = ""
        _meta["descricao"] = _meta["descricao"].fillna("")

        if not _est_ini_df.empty and "estoque_total" in _est_ini_df.columns:
            _meta = _meta.merge(
                _est_ini_df[["material", "estoque_total"]].rename(columns={"estoque_total": "EST.INI"}),
                on="material", how="left",
            )
            _meta["EST.INI"] = _meta["EST.INI"].fillna(0).round(0).astype(int)
        else:
            _meta["EST.INI"] = 0

        # ── Pivot "com pedidos" ───────────────────────────────────────────────
        _mes_cols_orig = sorted(df_mrp["mes"].unique())
        _proj_com = df_mrp.pivot_table(
            index="material", columns="mes", values="estoque_proj", aggfunc="sum"
        ).reset_index()

        # ── Pivot "sem pedidos" ───────────────────────────────────────────────
        if not _mrp_full.empty and "total_entradas" in _mrp_full.columns:
            _sp = _mrp_full.sort_values(["material", "periodo"]).copy()
            _sp["cum_ent"] = _sp.groupby("material")["total_entradas"].cumsum()
            _sp["est_sem"] = _sp["estoque_projetado"] - _sp["cum_ent"]
            _proj_sem = _sp.pivot_table(
                index="material", columns="periodo", values="est_sem", aggfunc="sum"
            ).reset_index().rename(columns={"material": "material"})
            _mes_cols_sem = sorted([c for c in _proj_sem.columns if c != "material"])
        else:
            _proj_sem = _proj_com.copy()
            _mes_cols_sem = _mes_cols_orig

        # Labels PT-BR para meses
        _mes_labels_com = [_mes_label(m) for m in _mes_cols_orig]
        _proj_com = _proj_com.rename(columns={o: _mes_label(o) for o in _mes_cols_orig})
        _mes_labels_sem = [_mes_label(m) for m in _mes_cols_sem]
        _proj_sem = _proj_sem.rename(columns={o: _mes_label(o) for o in _mes_cols_sem})

        # Meses com pedidos (entrada > 0)
        _ent_por_mat_mes: dict[str, set[str]] = _col_mod.defaultdict(set)
        if not _mrp_full.empty and "total_entradas" in _mrp_full.columns:
            for _, _er in _mrp_full[_mrp_full["total_entradas"] > 0].iterrows():
                _ent_por_mat_mes[str(_er["material"])].add(_mes_label(str(_er["periodo"])))

        # ── Tooltip: resumo de pedidos por material ───────────────────────────
        def _build_tooltip(mat: str) -> str:
            linhas = []
            if not _abertos.empty:
                sub = _abertos[_abertos["material"].astype(str) == str(mat)]
                for _, ro in sub.iterrows():
                    forn  = str(ro.get("fornecedor", "—"))
                    qtd   = ro.get("quantidade", 0)
                    mes_e = str(ro.get("mes_remessa", "—"))
                    ped   = str(ro.get("numero_pedido", "—"))
                    linhas.append(f"<tr><td>📦 Em aberto</td><td>{ped}</td>"
                                  f"<td>{forn[:22]}</td>"
                                  f"<td style='text-align:right'>{int(qtd):,}</td>"
                                  f"<td>{mes_e}</td></tr>")
            if not _novos_ped.empty:
                sub = _novos_ped[_novos_ped["material"].astype(str) == str(mat)]
                for _, ro in sub.iterrows():
                    qtd   = ro.get("quantidade", 0)
                    ent   = str(ro.get("periodo_entrega", "—"))
                    linhas.append(f"<tr><td style='color:#a8e6cf'>🔧 Sugerido MRP</td>"
                                  f"<td>—</td><td>—</td>"
                                  f"<td style='text-align:right'>{int(qtd):,}</td>"
                                  f"<td>{ent}</td></tr>")
            if not linhas:
                return ("<div style='color:#aaa;font-style:italic;padding:4px'>"
                        "Nenhum pedido em andamento</div>")
            rows = "".join(linhas[:10])
            extra = f"<tr><td colspan='5' style='color:#888;font-size:10px'>+{len(linhas)-10} mais…</td></tr>" if len(linhas) > 10 else ""
            return (f"<table style='width:100%;border-collapse:collapse;font-size:11px'>"
                    f"<tr style='color:#aaa;font-size:10px'>"
                    f"<th>Tipo</th><th>Pedido</th><th>Fornecedor</th>"
                    f"<th>Qtd</th><th>Entrega</th></tr>"
                    f"{rows}{extra}</table>")

        _tooltip_map = {str(m): _build_tooltip(str(m))
                        for m in _meta["material"].unique()}

        # ── Filtros ───────────────────────────────────────────────────────────
        st.markdown("### Filtros")
        _fr1, _fr2, _fr3 = st.columns([3, 3, 3])
        _fr4, _fr5, _fr6 = st.columns([3, 3, 3])

        _search_cod  = _fr1.text_input("Código do material", "", placeholder="Ex: 400040")
        _search_desc = _fr2.text_input("Descrição", "", placeholder="Buscar por nome...")
        _all_mes_opts = _mes_labels_com
        _sel_meses = _fr3.multiselect("Meses", _all_mes_opts, default=_all_mes_opts,
                                       placeholder="Selecione meses...")

        _deptos, _progs = [], []
        if not _dem_det.empty:
            if "departamento" in _dem_det.columns:
                _deptos = sorted(_dem_det["departamento"].dropna().unique())
            if "programa_orcamentario" in _dem_det.columns:
                _progs  = sorted(_dem_det["programa_orcamentario"].dropna().unique())

        _sel_depto = _fr4.multiselect("Departamento", _deptos,
                                       placeholder="Todos os departamentos")
        _sel_prog  = _fr5.multiselect("Programa Orçamentário", _progs,
                                       placeholder="Todos os programas")
        _sel_status = _fr6.selectbox(
            "Status de estoque",
            ["Todos", "Com ruptura", "Sem ruptura", "Abaixo do SS", "Com pedido", "Sem pedido"],
        )

        # Filtrar materiais por depto/prog via demanda_detail
        _mats_permitidos = set(_meta["material"].astype(str))
        if (_sel_depto or _sel_prog) and not _dem_det.empty:
            _dd = _dem_det.copy()
            if _sel_depto:
                _dd = _dd[_dd["departamento"].isin(_sel_depto)]
            if _sel_prog:
                _dd = _dd[_dd["programa_orcamentario"].isin(_sel_prog)]
            _mats_permitidos = set(_dd["material"].astype(str))

        # ── Tabela completa com metadados ─────────────────────────────────────
        _disp_com = _meta.merge(_proj_com, on="material", how="left").copy()
        _disp_sem = _meta.merge(_proj_sem, on="material", how="left").copy()

        def _aplicar_filtros(df_in, mes_labels):
            d = df_in.copy()
            d = d[d["material"].astype(str).isin(_mats_permitidos)]
            if _search_cod:
                d = d[d["material"].astype(str).str.contains(_search_cod, case=False, na=False)]
            if _search_desc:
                d = d[d["descricao"].astype(str).str.contains(_search_desc, case=False, na=False)]
            _mcols = [m for m in mes_labels if m in d.columns]
            if _sel_status == "Com ruptura":
                d = d[(d[_mcols].fillna(0) <= 0).any(axis=1)]
            elif _sel_status == "Sem ruptura":
                d = d[(d[_mcols].fillna(0) > 0).all(axis=1)]
            elif _sel_status == "Abaixo do SS":
                d = d[(d[_mcols].fillna(0).lt(d["EST.SEG."].values.reshape(-1, 1))).any(axis=1)]
            elif _sel_status == "Com pedido":
                _tem_ped = {m for m, s in _ent_por_mat_mes.items() if s}
                d = d[d["material"].astype(str).isin(_tem_ped)]
            elif _sel_status == "Sem pedido":
                _tem_ped = {m for m, s in _ent_por_mat_mes.items() if s}
                d = d[~d["material"].astype(str).isin(_tem_ped)]
            return d, _mcols

        _df_com, _mcols_com = _aplicar_filtros(_disp_com, _mes_labels_com)
        _df_sem, _mcols_sem = _aplicar_filtros(_disp_sem, _mes_labels_sem)

        # Meses selecionados como filtro adicional
        if _sel_meses:
            _mcols_com = [m for m in _mcols_com if m in _sel_meses]
            _mcols_sem = [m for m in _mcols_sem if m in _sel_meses]

        # ── KPIs ─────────────────────────────────────────────────────────────
        st.markdown("---")
        _k1,_k2,_k3,_k4,_k5,_k6,_k7 = st.columns(7)

        def _kpi_levels(df_k, mcols):
            if df_k.empty or not mcols:
                return 0, 0, 0, 0, 0, 0, 0
            _vals = df_k[mcols].fillna(0)
            _ss_v = df_k["EST.SEG."].values.reshape(-1, 1)
            n_tot  = len(df_k)
            _m_rupt = (_vals <= 0).any(axis=1)                       # tem ruptura
            _m_abx  = (~_m_rupt) & (_vals < _ss_v).any(axis=1)      # abaixo SS, sem ruptura
            _m_ok   = ~_m_rupt & ~_m_abx                             # tudo dentro do SS
            n_rupt  = int(_m_rupt.sum())
            n_abx   = int(_m_abx.sum())
            n_ok    = int(_m_ok.sum())
            _tem_ped = {m for m, s in _ent_por_mat_mes.items() if s}
            n_cpd   = int(df_k["material"].astype(str).isin(_tem_ped).sum())
            n_spd   = n_tot - n_cpd
            n_scob  = n_rupt   # sem cobertura = tem ruptura no período
            return n_tot, n_ok, n_abx, n_rupt, n_cpd, n_spd, n_scob

        _n_tot,_n_ok,_n_abx,_n_rupt,_n_cpd,_n_spd,_n_scob = _kpi_levels(_df_com, _mcols_com)
        _k1.metric("📦 Itens", _n_tot)
        _k2.metric("🟢 Adequado", _n_ok)
        _k3.metric("🟡 Alerta", _n_abx)
        _k4.metric("🔴 Ruptura", _n_rupt)
        _k5.metric("✅ Com pedido", _n_cpd)
        _k6.metric("⬜ Sem pedido", _n_spd)
        _k7.metric("⚠ Sem cobertura", _n_scob)
        st.markdown("---")

        # ── Função geradora de HTML da tabela com tooltip ─────────────────────
        def _render_proj_html(df_t, mcols, titulo):
            if df_t.empty:
                return f"<p style='color:#888'>{titulo}: nenhum dado.</p>"

            CSS = """
<style>
.pt-wrap{overflow-x:auto;overflow-y:auto;max-height:560px;border:1px solid #ddd;border-radius:6px}
.pt{border-collapse:collapse;width:100%;font-size:12px;font-family:'Segoe UI',sans-serif}
.pt thead tr{background:#1a4276;color:#fff;position:sticky;top:0;z-index:50}
.pt th{padding:6px 9px;text-align:right;white-space:nowrap;border:1px solid #0d2b54;font-size:11px}
.pt th.thl{text-align:left}
.pt td{padding:4px 8px;border:1px solid #e8e8e8;white-space:nowrap;vertical-align:middle}
.pt tr:nth-child(even){background:#fafafa}
.pt tr:hover{background:#eaf3ff!important}
.cls-a{background:#c0392b;color:#fff;font-weight:700;text-align:center;border-radius:3px;padding:2px 5px}
.cls-b{background:#e67e22;color:#fff;font-weight:700;text-align:center;border-radius:3px;padding:2px 5px}
.cls-c{background:#2980b9;color:#fff;font-weight:700;text-align:center;border-radius:3px;padding:2px 5px}
.ss-col{background:#90EE90!important;font-weight:700;text-align:right}
.mx-col{background:#d4f1c4!important;text-align:right}
.num{text-align:right}
.st-r{background:#e74c3c;color:#fff;font-weight:700;text-align:right}
.st-o{background:#e67e22;color:#fff;text-align:right}
.st-y{background:#f9e04b;color:#333;text-align:right}
.st-g{background:#b7e1b0;color:#222;text-align:right}
.th-ss{background:#2ecc71!important}
.th-mx{background:#27ae60!important}
/* Tooltip */
.tip-host{position:relative;cursor:help}
.tip-host .tip{
  display:none;position:absolute;left:0;bottom:110%;z-index:99999;
  background:#1a1a2e;color:#f0f0f0;border-radius:8px;padding:12px 14px;
  min-width:320px;max-width:420px;box-shadow:0 6px 20px rgba(0,0,0,.55);
  border:1px solid #444;pointer-events:none;white-space:normal;
  font-size:11px;line-height:1.4
}
.tip-host:hover .tip{display:block}
.tip-title{font-size:13px;font-weight:700;color:#7fc7ff;margin-bottom:8px}
.tip table{width:100%;border-collapse:collapse}
.tip th{color:#aaa;font-size:10px;text-align:left;padding:2px 4px;border-bottom:1px solid #333}
.tip td{padding:3px 5px;color:#fff;border-bottom:1px solid #2a2a3a}
.tip tr:last-child td{border-bottom:none}
</style>
"""
            def _fmt(v):
                if pd.isna(v): return "—"
                return f"{int(round(v)):,}".replace(",",".")

            def _cls_cell(cls):
                c = {"A":"cls-a","B":"cls-b","C":"cls-c"}.get(str(cls),"")
                return f'<span class="{c}">{cls}</span>'

            def _st_cell(v, ss):
                if pd.isna(v): return '<td class="num" style="color:#bbb">—</td>'
                vi = int(round(v))
                f  = _fmt(v)
                if vi <= 0:   return f'<td class="st-r">{f}</td>'
                if vi < ss:   return f'<td class="st-o">{f}</td>'
                if vi < ss*1.15: return f'<td class="st-y">{f}</td>'
                return f'<td class="st-g">{f}</td>'

            headers = (
                '<th class="thl" style="min-width:36px">CLS</th>'
                '<th class="thl" style="min-width:75px">CÓDIGO</th>'
                '<th class="thl" style="min-width:180px">DESCRIÇÃO</th>'
                '<th style="min-width:65px">EST.INI</th>'
                '<th class="th-ss" style="min-width:65px">EST.SEG.</th>'
                '<th class="th-mx" style="min-width:65px">EST.MÁX.</th>'
                '<th style="min-width:55px">LEAD(D)</th>'
            )
            for m in mcols:
                headers += f'<th style="min-width:70px">{m}</th>'

            rows_html = []
            for _, row in df_t.iterrows():
                mat = str(row["material"])
                ss  = int(row.get("EST.SEG.", 0) or 0)
                tip_content = _tooltip_map.get(mat, "Sem pedidos registrados")
                tip_html = (
                    f'<div class="tip">'
                    f'<div class="tip-title">📦 Pedidos — {mat}</div>'
                    f'{tip_content}</div>'
                )
                cells = (
                    f'<td>{_cls_cell(row.get("classe",""))}</td>'
                    f'<td class="tip-host">{mat}{tip_html}</td>'
                    f'<td>{str(row.get("descricao",""))[:42]}</td>'
                    f'<td class="num">{_fmt(row.get("EST.INI",0))}</td>'
                    f'<td class="ss-col">{_fmt(row.get("EST.SEG.",0))}</td>'
                    f'<td class="mx-col">{_fmt(row.get("EST.MÁX.",0))}</td>'
                    f'<td class="num">{int(row.get("LEAD(D)", LEAD_TIME_DIAS))}</td>'
                )
                for m in mcols:
                    cells += _st_cell(row.get(m), ss)
                rows_html.append(f"<tr>{cells}</tr>")

            n_rows = len(rows_html)
            height = min(max(200, 50 + n_rows * 34), 600)
            html = (
                f'{CSS}<div class="pt-wrap" style="height:{height}px">'
                f'<table class="pt"><thead><tr>{headers}</tr></thead>'
                f'<tbody>{"".join(rows_html)}</tbody></table></div>'
            )
            return html

        # ── Subtabs: Com Pedidos / Sem Pedidos ────────────────────────────────
        st.info(
            "Células **vermelhas** = ruptura esperada durante o trânsito do pedido. "
            "Passe o mouse sobre o **código** do material para ver os pedidos em andamento.",
            icon="ℹ️",
        )
        _stab_com, _stab_sem = st.tabs([
            "📦 Com Pedidos (realizados + MRP)",
            "📉 Sem Pedidos (consumo puro)",
        ])
        with _stab_com:
            _stc.html(_render_proj_html(_df_com, _mcols_com, "Com Pedidos"), height=640, scrolling=True)
        with _stab_sem:
            _stc.html(_render_proj_html(_df_sem, _mcols_sem, "Sem Pedidos"), height=640, scrolling=True)

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

        # ── Raio-X de Auditoria (sempre visível, antes dos filtros) ──────────
        with st.expander("🕵️ Raio-X de Auditoria Financeira (Buscando Divergências)", expanded=False):

            # ── Totais brutos absolutos (sem filtro algum) ────────────────────
            _rx_hist  = df_mb51["valor_pedido"].sum()       if not df_mb51.empty        and "valor_pedido"       in df_mb51.columns        else 0.0
            _rx_exist = df_abertos_fut["valor_total_pedido"].sum() if not df_abertos_fut.empty and "valor_total_pedido" in df_abertos_fut.columns else 0.0
            _rx_novo  = df_ped["valor_total_pedido"].sum()  if not df_ped.empty         and "valor_total_pedido" in df_ped.columns          else 0.0

            _rx1, _rx2, _rx3 = st.columns(3)
            _rx1.metric("📦 Histórico recebido (MB51)",      _fmt_brl_contabil(_rx_hist))
            _rx2.metric("🔄 Pedidos existentes (SAP)",       _fmt_brl_contabil(_rx_exist))
            _rx3.metric("🛒 Novos pedidos MRP",              _fmt_brl_contabil(_rx_novo))

            st.caption(
                f"Linhas em memória — "
                f"MB51: **{len(df_mb51):,}** · "
                f"Pedidos SAP: **{len(df_abertos_fut):,}** · "
                f"Pedidos MRP: **{len(df_ped):,}** "
                f"(compare com a contagem de linhas dos seus arquivos CSV)"
            )

            st.divider()

            # ── Caça ao Preço Zero ────────────────────────────────────────────
            _preco_zero: list[pd.DataFrame] = []

            if not df_ped.empty:
                _mask_ped = (df_ped["quantidade"] > 0) & (
                    (df_ped.get("valor_unitario",    pd.Series(dtype=float)).fillna(0) == 0) |
                    (df_ped.get("valor_total_pedido", pd.Series(dtype=float)).fillna(0) == 0)
                )
                _pz_ped = df_ped.loc[_mask_ped, ["material"]].copy()
                _pz_ped["quantidade"]   = df_ped.loc[_mask_ped, "quantidade"]
                _pz_ped["origem"]       = "Novo Pedido (MRP)"
                _preco_zero.append(_pz_ped)

            if not df_abertos_fut.empty:
                _col_val_ped = "valor_total_pedido" if "valor_total_pedido" in df_abertos_fut.columns else None
                if _col_val_ped:
                    _mask_ex = (df_abertos_fut["quantidade"] > 0) & (
                        df_abertos_fut[_col_val_ped].fillna(0) == 0
                    )
                    _pz_ex = df_abertos_fut.loc[_mask_ex, ["material"]].copy()
                    _pz_ex["quantidade"] = df_abertos_fut.loc[_mask_ex, "quantidade"]
                    _pz_ex["origem"]     = "Pedido Existente (SAP)"
                    _preco_zero.append(_pz_ex)

            if _preco_zero:
                _df_pz = (
                    pd.concat(_preco_zero, ignore_index=True)
                    .groupby(["material", "origem"], as_index=False)["quantidade"]
                    .sum()
                    .sort_values("quantidade", ascending=False)
                    .head(50)
                    .reset_index(drop=True)
                )
            else:
                _df_pz = pd.DataFrame()

            if not _df_pz.empty:
                st.warning(
                    "⚠️ Atenção: Os materiais abaixo possuem quantidade a receber/comprar, "
                    "mas o valor financeiro está R$ 0,00 "
                    "(Falta preço no cadastro ou contrato)."
                )
                st.dataframe(_df_pz, use_container_width=True)
            else:
                st.success("✅ Nenhum item com quantidade > 0 e preço zerado encontrado.")

        if df_fin_bruto.empty and df_fluxo_bruto.empty:
            st.info("Nenhum dado financeiro disponível.")
        else:
            # ── Filtros globais ───────────────────────────────────────────────
            _ano_atual = str(pd.Timestamp.now().year)

            # Coletar anos disponíveis em ambas as visões
            _anos_orc = (
                df_fin_bruto["mes_pedido"]
                .dropna()
                .str[:4]
                .unique()
                .tolist()
                if not df_fin_bruto.empty and "mes_pedido" in df_fin_bruto.columns
                else []
            )
            _anos_cx = (
                df_fluxo_bruto["mes_pagamento"]
                .dropna()
                .str[:4]
                .unique()
                .tolist()
                if not df_fluxo_bruto.empty and "mes_pagamento" in df_fluxo_bruto.columns
                else []
            )
            _anos_disp = sorted(set(_anos_orc + _anos_cx))
            _default_ano = _ano_atual if _ano_atual in _anos_disp else (_anos_disp[0] if _anos_disp else _ano_atual)

            _deptos_disp = sorted(
                df_fin_bruto["departamento"].dropna().unique().tolist()
                if not df_fin_bruto.empty and "departamento" in df_fin_bruto.columns
                else []
            )
            _progs_disp  = sorted(
                df_fin_bruto["programa_orcamentario"].dropna().unique().tolist()
                if not df_fin_bruto.empty and "programa_orcamentario" in df_fin_bruto.columns
                else []
            )

            _fc1, _fc2, _fc3 = st.columns(3)
            with _fc1:
                _sel_ano   = st.selectbox("Ano", _anos_disp,
                                          index=_anos_disp.index(_default_ano) if _default_ano in _anos_disp else 0,
                                          key="fin_ano")
            with _fc2:
                _sel_depto = st.selectbox("Departamento", ["Todos"] + _deptos_disp, key="fin_depto")
            with _fc3:
                _sel_prog  = st.selectbox("Programa Orçamentário", ["Todos"] + _progs_disp, key="fin_prog")

            # ── Filtrar brutos ────────────────────────────────────────────────
            def _filtrar_fin(df: pd.DataFrame, col_mes: str) -> pd.DataFrame:
                if df.empty:
                    return df
                out = df[df[col_mes].str[:4] == _sel_ano].copy() if col_mes in df.columns else df.copy()
                if _sel_depto != "Todos" and "departamento" in out.columns:
                    out = out[out["departamento"] == _sel_depto]
                if _sel_prog  != "Todos" and "programa_orcamentario" in out.columns:
                    out = out[out["programa_orcamentario"] == _sel_prog]
                return out

            _fin_f   = _filtrar_fin(df_fin_bruto,   "mes_pedido")
            _fluxo_f = _filtrar_fin(df_fluxo_bruto, "mes_pagamento")

            # ── Reconstruir pivots dinâmicos ──────────────────────────────────
            def _pivot_com_total(df: pd.DataFrame, idx: str, col_val: str) -> pd.DataFrame:
                if df.empty or idx not in df.columns or "origem" not in df.columns:
                    return pd.DataFrame()
                pv = (
                    df.groupby([idx, "origem"], as_index=False)[col_val]
                    .sum()
                    .pivot(index=idx, columns="origem", values=col_val)
                    .fillna(0)
                    .sort_index()
                )
                pv.columns.name = None
                pv["Total"] = pv.sum(axis=1)
                total_row = pv.sum(numeric_only=True)
                total_row.name = "TOTAL GERAL"
                pv = pd.concat([pv, total_row.to_frame().T])
                return pv

            _vis_orc_f = _pivot_com_total(_fin_f,   "mes_pedido",    "valor_rateado")
            _vis_cx_f  = _pivot_com_total(_fluxo_f, "mes_pagamento", "valor_rateado")

            subtab_orc, subtab_cx = st.tabs([
                "📊 Visão Orçamentária (Emissão)",
                "💸 Visão de Caixa (Desembolso Real)",
            ])

            with subtab_orc:
                st.caption(
                    f"Compromisso financeiro · {_sel_ano} · "
                    + (f"Departamento: {_sel_depto} · " if _sel_depto != "Todos" else "")
                    + (f"Programa: {_sel_prog}" if _sel_prog != "Todos" else "Todos os departamentos/programas")
                )
                if _vis_orc_f.empty:
                    st.info("Sem dados orçamentários para os filtros selecionados.")
                else:
                    _chart_financeiro(_vis_orc_f.drop("TOTAL GERAL", errors="ignore"), "Compromisso por Mês de Emissão")
                    st.caption("💡 Clique em uma linha para ver o detalhamento dos pedidos daquele mês.")
                    agg_fmt = _vis_orc_f.copy()
                    for _c in agg_fmt.columns:
                        agg_fmt[_c] = agg_fmt[_c].apply(_fmt_brl_contabil)
                    _sel_orc = st.dataframe(
                        agg_fmt,
                        use_container_width=True,
                        on_select="rerun",
                        selection_mode="single-row",
                        key="orc_tbl_sel",
                    )

                    # ── Drill-down Orçamentário ───────────────────────────────
                    _orc_row_idxs = _sel_orc.selection.rows if hasattr(_sel_orc, "selection") else []
                    if _orc_row_idxs:
                        _orc_mes = list(_vis_orc_f.index)[_orc_row_idxs[0]]
                        if _orc_mes == "TOTAL GERAL":
                            st.info("Selecione um mês específico (não o total) para ver o detalhamento.")
                        else:
                            st.markdown(f"#### 🔍 Detalhes · Emissão do Pedido: **{_orc_mes}**")
                            _dd_orig_orc_opts = (
                                ["Todas"] + sorted(_fin_f["origem"].dropna().unique().tolist())
                                if not _fin_f.empty and "origem" in _fin_f.columns else ["Todas"]
                            )
                            _dd_orig_orc = st.selectbox("Filtrar por Origem", _dd_orig_orc_opts, key="orc_dd_orig")

                            _detail_orc = _fin_f[_fin_f["mes_pedido"] == _orc_mes].copy()
                            if _dd_orig_orc != "Todas":
                                _detail_orc = _detail_orc[_detail_orc["origem"] == _dd_orig_orc]

                            if _detail_orc.empty:
                                st.info("Nenhum registro para este mês/origem.")
                            else:
                                _orc_dcols = [c for c in [
                                    "material", "origem",
                                    "mes_pedido", "mes_entrega",
                                    "departamento", "programa_orcamentario",
                                    "valor_rateado",
                                ] if c in _detail_orc.columns]
                                _detail_orc_show = (
                                    _detail_orc[_orc_dcols]
                                    .rename(columns={
                                        "material"              : "Material",
                                        "origem"                : "Origem",
                                        "mes_pedido"            : "Mês Emissão",
                                        "mes_entrega"           : "Mês Chegada",
                                        "departamento"          : "Departamento",
                                        "programa_orcamentario" : "Programa",
                                        "valor_rateado"         : "Valor Rateado",
                                    })
                                    .sort_values(["Mês Chegada", "Material"])
                                    .reset_index(drop=True)
                                )
                                _tot_orc = _detail_orc_show["Valor Rateado"].sum()
                                st.caption(
                                    f"{len(_detail_orc_show)} linha(s) · "
                                    f"Total: **{_fmt_brl_contabil(_tot_orc)}**"
                                )
                                st.dataframe(
                                    _detail_orc_show.style.format(
                                        {"Valor Rateado": _fmt_brl_contabil}, na_rep="-"
                                    ),
                                    use_container_width=True,
                                    height=380,
                                )

            with subtab_cx:
                st.caption(
                    f"Desembolso previsto · {_sel_ano} · política padrão: 50% em 60 dias + 50% em 90 dias · "
                    + (f"Departamento: {_sel_depto} · " if _sel_depto != "Todos" else "")
                    + (f"Programa: {_sel_prog}" if _sel_prog != "Todos" else "Todos os departamentos/programas")
                )
                if _vis_cx_f.empty:
                    st.info("Nenhuma parcela de pagamento para os filtros selecionados.")
                else:
                    _chart_financeiro(_vis_cx_f.drop("TOTAL GERAL", errors="ignore"), "Desembolso por Mês de Pagamento")
                    st.caption("💡 Clique em uma linha para rastrear de onde vem o desembolso daquele mês.")
                    agg_cx_fmt = _vis_cx_f.copy()
                    for _c in agg_cx_fmt.columns:
                        agg_cx_fmt[_c] = agg_cx_fmt[_c].apply(_fmt_brl_contabil)
                    _sel_cx = st.dataframe(
                        agg_cx_fmt,
                        use_container_width=True,
                        on_select="rerun",
                        selection_mode="single-row",
                        key="cx_tbl_sel",
                    )

                    # ── Drill-down Caixa ──────────────────────────────────────
                    _cx_row_idxs = _sel_cx.selection.rows if hasattr(_sel_cx, "selection") else []
                    if _cx_row_idxs:
                        _cx_mes = list(_vis_cx_f.index)[_cx_row_idxs[0]]
                        if _cx_mes == "TOTAL GERAL":
                            st.info("Selecione um mês específico (não o total) para ver o detalhamento.")
                        else:
                            st.markdown(f"#### 🔍 Detalhes · Desembolso em: **{_cx_mes}**")
                            _dd_orig_cx_opts = (
                                ["Todas"] + sorted(_fluxo_f["origem"].dropna().unique().tolist())
                                if not _fluxo_f.empty and "origem" in _fluxo_f.columns else ["Todas"]
                            )
                            _dd_orig_cx = st.selectbox("Filtrar por Origem", _dd_orig_cx_opts, key="cx_dd_orig")

                            _detail_cx = _fluxo_f[_fluxo_f["mes_pagamento"] == _cx_mes].copy()
                            if _dd_orig_cx != "Todas":
                                _detail_cx = _detail_cx[_detail_cx["origem"] == _dd_orig_cx]

                            if _detail_cx.empty:
                                st.info("Nenhum registro para este mês/origem.")
                            else:
                                _cx_dcols = [c for c in [
                                    "material", "origem",
                                    "mes_emissao",   # data geração pedido
                                    "mes_entrega",   # data chegada do material
                                    "prazo_dias",    # prazo de pagamento
                                    "mes_pagamento", # data desembolso
                                    "departamento", "programa_orcamentario",
                                    "valor_rateado",
                                ] if c in _detail_cx.columns]
                                _detail_cx_show = (
                                    _detail_cx[_cx_dcols]
                                    .rename(columns={
                                        "material"              : "Material",
                                        "origem"                : "Origem",
                                        "mes_emissao"           : "Data Emissão PO",
                                        "mes_entrega"           : "Data Chegada",
                                        "prazo_dias"            : "Prazo Pgto (dias)",
                                        "mes_pagamento"         : "Data Desembolso",
                                        "departamento"          : "Departamento",
                                        "programa_orcamentario" : "Programa",
                                        "valor_rateado"         : "Valor Desembolso",
                                    })
                                    .sort_values(["Data Chegada", "Material"])
                                    .reset_index(drop=True)
                                )
                                _tot_cx = _detail_cx_show["Valor Desembolso"].sum()
                                st.caption(
                                    f"{len(_detail_cx_show)} parcela(s) · "
                                    f"Total: **{_fmt_brl_contabil(_tot_cx)}**"
                                )
                                st.dataframe(
                                    _detail_cx_show.style.format(
                                        {"Valor Desembolso": _fmt_brl_contabil}, na_rep="-"
                                    ),
                                    use_container_width=True,
                                    height=380,
                                )

                    # ── Rastreio completo: Emissão → Entrega → Pagamento ──────
                    with st.expander("🔍 Rastreio: Emissão → Entrega → Pagamento"):
                        st.caption(
                            "Cada linha representa uma parcela de pagamento rateada. "
                            "Rastreie pela coluna **Mês Emissão** (quando o PO foi criado) "
                            "e **Mês Entrega** (quando o material chega ao estoque)."
                        )
                        if not _fluxo_f.empty:
                            _trace_cols = [c for c in [
                                "origem", "material", "departamento", "programa_orcamentario",
                                "mes_emissao", "mes_entrega", "prazo_dias", "mes_pagamento", "valor_rateado",
                            ] if c in _fluxo_f.columns]
                            df_trace = (
                                _fluxo_f[_trace_cols]
                                .rename(columns={
                                    "origem"                : "Origem",
                                    "material"              : "Material",
                                    "departamento"          : "Departamento",
                                    "programa_orcamentario" : "Programa",
                                    "mes_emissao"           : "Mês Emissão",
                                    "mes_entrega"           : "Mês Entrega",
                                    "prazo_dias"            : "Prazo (dias)",
                                    "mes_pagamento"         : "Mês Pagamento",
                                    "valor_rateado"         : "Valor Rateado",
                                })
                                .sort_values(["Mês Pagamento", "Mês Entrega"])
                                .reset_index(drop=True)
                            )
                            st.dataframe(
                                df_trace.style.format(
                                    {"Valor Rateado": _fmt_brl_contabil}, na_rep="-"
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
