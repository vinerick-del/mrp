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
    achar_arquivo,
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
    "pcm"       : "pcm_materiais.xlsx",
}


def _path_arquivo(chave: str) -> str:
    """Resolve caminho case-insensitive: usa achar_arquivo para tolerar
    variações de capitalização (Demanda.csv, REMESSAS_SAP.csv, etc.)."""
    nome    = _ARQUIVOS_MAPA[chave]
    resolvido = achar_arquivo(nome)
    return resolvido if resolvido else os.path.join(DIR_DADOS, nome)


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
# HELPERS DE RATEIO MANUAL
# ─────────────────────────────────────────────────────────────────────────────
_RATEIO_MANUAL_PATH = os.path.join(DIR_DADOS, "rateio_manual.csv")


def _carregar_rateio_manual() -> pd.DataFrame:
    """Retorna DataFrame do rateio_manual.csv ou vazio se não existir."""
    if not os.path.exists(_RATEIO_MANUAL_PATH):
        return pd.DataFrame(
            columns=["material", "departamento", "programa_orcamentario", "proporcao", "atualizado_em"]
        )
    try:
        return pd.read_csv(_RATEIO_MANUAL_PATH, sep=";", dtype={"material": str})
    except Exception:
        return pd.DataFrame(
            columns=["material", "departamento", "programa_orcamentario", "proporcao", "atualizado_em"]
        )


def _salvar_rateio_manual(material: str, linhas: list[dict]) -> None:
    """Persiste rateio manual para um material.

    Args:
        material: código do material.
        linhas: lista de dicts com chaves departamento, programa_orcamentario,
                proporcao_pct (valor 0-100; será convertido para 0-1 ao salvar).
    """
    df = _carregar_rateio_manual()
    df = df[df["material"].astype(str) != str(material)].copy()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = pd.DataFrame([
        {
            "material"              : str(material),
            "departamento"          : l["departamento"],
            "programa_orcamentario" : l["programa_orcamentario"],
            "proporcao"             : round(l["proporcao_pct"] / 100.0, 6),
            "atualizado_em"         : now_str,
        }
        for l in linhas
    ])
    df = pd.concat([df, new_rows], ignore_index=True)
    os.makedirs(DIR_DADOS, exist_ok=True)
    df.to_csv(_RATEIO_MANUAL_PATH, sep=";", index=False)


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
# HELPER: classificar pedidos MRP por cobertura de contrato
# ─────────────────────────────────────────────────────────────────────────────
def _classificar_contratos_mrp(
    df_ped: pd.DataFrame, contratos: pd.DataFrame
) -> pd.DataFrame:
    """
    Classifica cada pedido MRP gerado como coberto ou não por contrato vigente.

    Regras:
      saldo_contrato == 1  → Acordo de Preço  : quantidade ilimitada;
                              coberto se data_fim_vigencia >= data_pedido.
      saldo_contrato >  1  → Acordo de Quantidade: saldo deduzido cronologicamente.
      Sem contrato vigente → "Sem Contrato".

    Pode expandir linhas quando um pedido é parcialmente coberto por saldo.

    Retorna DataFrame com colunas adicionais:
      _origem_mrp      — "Novo Pedido (MRP) — Com Contrato" / "… — Sem Contrato"
      _status_contrato — detalhe (Com Contrato / Contrato Vencido / Saldo Esgotado …)
      _tipo_contrato   — "Acordo de Preço" / "Acordo de Quantidade" / "—"
    """
    if df_ped.empty:
        out = df_ped.copy()
        for _c in ["_origem_mrp", "_status_contrato", "_tipo_contrato"]:
            out[_c] = ""
        return out

    # Mapa: material → {tipo, saldo, vigencia}
    _cmap: dict = {}
    if not contratos.empty and "saldo_contrato" in contratos.columns:
        for _, _cr in contratos.iterrows():
            _mat_c = str(_cr["material"]).strip()
            _saldo_c = float(_cr.get("saldo_contrato") or 0)
            _vig_raw = _cr.get("data_fim_vigencia")
            _vig_ts = None
            if pd.notna(_vig_raw):
                _vig_ts = pd.to_datetime(_vig_raw, format="%d/%m/%Y", errors="coerce")
                if pd.isna(_vig_ts):
                    _vig_ts = pd.to_datetime(_vig_raw, errors="coerce")
            _cmap[_mat_c] = {
                "tipo"    : "preco" if _saldo_c <= 1 else "quantidade",
                "saldo"   : _saldo_c,
                "vigencia": _vig_ts,
            }

    # Saldo corrente por material (deduzido cronologicamente para acordos de qtd)
    _saldo_rest: dict = {_m: _d["saldo"] for _m, _d in _cmap.items()}

    _df = df_ped.copy()
    _df["_dt_ped"] = pd.to_datetime(_df["data_pedido"], format="%d/%m/%Y", errors="coerce")
    _df = _df.sort_values("_dt_ped").reset_index(drop=True)

    _rows: list[dict] = []
    for _, _row in _df.iterrows():
        _mat = str(_row.get("material", "")).strip()
        _qty = float(_row.get("quantidade") or 0)
        _vu  = float(_row.get("valor_unitario") or 0)
        _dt  = _row["_dt_ped"]

        def _push(_origem, _qtd, _status, _tipo):
            _r = {**_row.to_dict(),
                  "quantidade"         : _qtd,
                  "valor_total_pedido" : _qtd * _vu,
                  "_origem_mrp"        : _origem,
                  "_status_contrato"   : _status,
                  "_tipo_contrato"     : _tipo}
            _rows.append(_r)

        if _mat not in _cmap:
            _push("Novo Pedido (MRP) — Sem Contrato", _qty, "Sem Contrato", "—")
        else:
            _ci  = _cmap[_mat]
            _vig = _ci["vigencia"]
            _vig_ok = (
                _vig is None or pd.isna(_vig)
                or (pd.notna(_dt) and _dt <= _vig)
            )
            _tl = "Acordo de Preço" if _ci["tipo"] == "preco" else "Acordo de Quantidade"

            if _ci["tipo"] == "preco":
                if _vig_ok:
                    _push("Novo Pedido (MRP) — Com Contrato",   _qty, "Com Contrato",    _tl)
                else:
                    _push("Novo Pedido (MRP) — Sem Contrato",   _qty, "Contrato Vencido", _tl)
            else:
                _bal = _saldo_rest.get(_mat, 0.0)
                if _bal <= 0:
                    _push("Novo Pedido (MRP) — Sem Contrato",   _qty, "Saldo Esgotado",  _tl)
                elif _bal >= _qty:
                    _saldo_rest[_mat] -= _qty
                    _push("Novo Pedido (MRP) — Com Contrato",   _qty, "Com Contrato",    _tl)
                else:
                    # Cobertura parcial → duas linhas
                    _saldo_rest[_mat] = 0.0
                    _push("Novo Pedido (MRP) — Com Contrato",   _bal,        "Com Contrato (parcial)", _tl)
                    _push("Novo Pedido (MRP) — Sem Contrato",   _qty - _bal, "Saldo Esgotado",         _tl)

    _out = pd.DataFrame(_rows)
    if "_dt_ped" in _out.columns:
        _out = _out.drop(columns=["_dt_ped"])
    return _out


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
    st.header("🚀 MRP — Processamento")
    btn_processar = st.button("🚀 Processar MRP", type="primary", use_container_width=True)
    if _tem_dados_minimos():
        st.caption("✅ Dados carregados — processamento automático ativo.")

    st.divider()

    # ── Menu de navegação vertical ────────────────────────────────────────────
    _menu_opcoes = ["📂 Importação de Arquivos"]
    if "resultado" in st.session_state:
        _menu_opcoes += [
            "📅 Projeção de Estoque",
            "💰 Financeiro",
            "📋 Saldo de Contrato",
            "📂 Rateio",
            "⚠️ Rateio Pendente",
            "📊 Rateio DFP - Realizado",
            "🎯 PCM — Materiais Críticos",
        ]

    _pagina = st.radio(
        "Navegação",
        _menu_opcoes,
        label_visibility="collapsed",
    )

    st.divider()

    if st.button("🔄 Forçar Recarregamento (Limpar Cache)", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # Indicador de rateio manual configurado
    _rm_sidebar = _carregar_rateio_manual()
    if not _rm_sidebar.empty and "material" in _rm_sidebar.columns:
        _n_rm = _rm_sidebar["material"].nunique()
        if _n_rm > 0:
            st.info(f"🔧 Rateio manual: **{_n_rm}** material(is)")

# ─────────────────────────────────────────────────────────────────────────────
# DEFINIÇÕES DE ARQUIVO PARA IMPORTAÇÃO
# ─────────────────────────────────────────────────────────────────────────────
_file_defs = [
    ("demanda",   "①", "Demanda (DTM)",                  ["csv","txt"], "demanda_dtm_raw.csv — separado por ';'"),
    ("remessas",  "②", "Remessas SAP",                   ["csv","txt"], "ME2M/ME9F — lookup de datas e nºs de documento"),
    ("pedidos",   "③", "Pedidos em Aberto",              ["csv","txt"], "pedidos_abertos.csv — base principal"),
    ("estoque",   "④", "Estoque SAP",                    ["csv","txt"], "MB52/MMBE — separado por TAB"),
    ("contratos", "⑤", "Contratos SAP",                  ["csv","txt"], "ME3M/ME3N — separado por TAB"),
    ("materiais", "⑥", "Materiais (catálogo)",           ["csv","txt"], "MM60/MM03 — CÓDIGO | DESCRIÇÃO | VALOR"),
    ("lead",      "⑦", "Lead Times",                     ["csv"],       "CSV: material,lead_time_dias"),
    ("mb51",      "⑧", "Histórico MB51",                 ["csv","txt"], "MB51 — movimentos 101/102"),
    ("politica",  "⑨", "Política de Pagamento",          ["csv","txt"], "CSV: documento | dias_parcela_1 | ..."),
    ("pcm",       "⑩", "PCM — Materiais Críticos",       ["xlsx"],      "Arquivo PCM_*.xlsx do sistema de materiais críticos"),
]

# Criar file_uploader para cada arquivo (será usado em Importação)
_file_uploaders = {}
for chave, numero, descricao, tipos, ajuda in _file_defs:
    _file_uploaders[chave] = None  # Será preenchido na aba de Importação


# ─────────────────────────────────────────────────────────────────────────────
# ROTEAMENTO DE PÁGINAS
# ─────────────────────────────────────────────────────────────────────────────
if _pagina == "📂 Importação de Arquivos":
    # Mostrar aba de importação quando não há dados processados
    st.markdown("## 📂 Importação de Arquivos")
    st.info(
        "👇 Selecione e carregue seus arquivos SAP abaixo. Os arquivos são salvos automaticamente. "
        "Na próxima abertura, o app carrega os dados da última importação."
    )

    cols_upload = st.columns(2)
    for idx, (chave, numero, descricao, tipos, ajuda) in enumerate(_file_defs):
        col = cols_upload[idx % 2]
        with col:
            _file_uploaders[chave] = st.file_uploader(
                _label_upload(numero, descricao, chave),
                type=tipos,
                key=f"up_{chave}",
                help=ajuda,
            )
            _caption_arquivo(chave)

    st.divider()
    st.subheader("📋 Formato esperado de cada arquivo")
    with st.expander("Ver especificações"):
        st.markdown("""
| # | Arquivo | Separador | Colunas-chave |
|---|---------|-----------|---|
| ① | Demanda DTM | `;` | CÓDIGO, MÊS, DEP., PROJETO, QTD |
| ② | Remessas SAP | `TAB` | Material, Data de remessa, a ser fornecida (quantidade) |
| ③ | Pedidos em Aberto | `TAB` ou `;` | Material, Quantidade, Valor total |
| ④ | Estoque SAP | `TAB` | Produto, Qtd.disponível |
| ⑤ | Contratos SAP | `TAB` | Material, Fim da validade, Qtd.prev.pendente, Preço líquido |
| ⑥ | Materiais | `;` ou `TAB` | CÓDIGO, DESCRIÇÃO, VALOR UNITÁRIO |
| ⑦ | Lead Times | `,` | material, lead_time_dias |
| ⑧ | Histórico MB51 | `TAB` | Material, Data, Quantidade |
| ⑨ | Política de Pagamento | `;` | documento, dias_parcela_1, dias_parcela_2 ... |
| ⑩ | PCM — Materiais Críticos | XLSX | Definido no sistema |

**Números em formato brasileiro:** `3.515,50` = 3515.50
**Datas:** `dd/mm/yyyy`
        """)

    # Processar uploads
    _novos_uploads = [k for k, v in _file_uploaders.items() if v is not None]
    if _novos_uploads:
        for chave, uploaded in _file_uploaders.items():
            if uploaded:
                _salvar_upload(uploaded, chave)
        # Forçar reprocessamento quando novos arquivos chegarem
        st.session_state.pop("resultado", None)
        st.session_state.pop("_auto_processado", None)
        st.rerun()

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

            # ── Contratos SAP (opcional) ──────────────────────────────────────
            cont_path = achar_arquivo("Contratos_SAP") or achar_arquivo("contratos_sap.csv")
            contratos = pd.DataFrame()
            if cont_path:
                try:
                    contratos = ler_contratos_sap(cont_path)
                    if not contratos.empty and "saldo_contrato" in contratos.columns:
                        contratos["tipo_contrato"] = contratos["saldo_contrato"].apply(
                            lambda _s: "Acordo de Preço" if float(_s or 0) <= 1 else "Acordo de Quantidade"
                        )
                except Exception as _e_cont:
                    print(f"  ⚠ Contratos SAP não carregado ({_e_cont}) — sem verificação de saldo.")

            # ── Lead Times ────────────────────────────────────────────────────
            # Prioridade: 1) coluna LT em materiais.csv  2) arquivo LT (case-insensitive)
            if "lead_time_dias" in materiais.columns:
                lt_dict = {
                    str(row["material"]): int(row["lead_time_dias"])
                    for _, row in materiais.iterrows()
                    if pd.notna(row["lead_time_dias"]) and row["lead_time_dias"] > 0
                }
                print(f"  [LT] Lead times de materiais.csv: {len(lt_dict)} itens")
            else:
                lt_path = achar_arquivo("LEAD_TIMES.csv") or achar_arquivo("lead_times.csv")
                lt_dict = ler_lead_times(lt_path) if lt_path else {}
                if not lt_dict:
                    print(f"  [LT] ⚠ Nenhum lead time carregado — usando default {LEAD_TIME_DIAS}d para todos")
                    print(f"  [LT]   Colunas em materiais.csv: {list(materiais.columns)}")

            # ── Pipeline MRP ──────────────────────────────────────────────────
            # passo_1_2_demanda() retorna (demanda, df_raw) em versões recentes
            # ou apenas demanda nas versões anteriores — compatível com ambas.
            _p12_result = passo_1_2_demanda()
            if isinstance(_p12_result, tuple):
                demanda, _demanda_detail_raw = _p12_result
            else:
                demanda = _p12_result
                raw_path = os.path.join(DIR_DADOS, ARQUIVO_DEMANDA_RAW) if ARQUIVO_DEMANDA_RAW else None
                _demanda_detail_raw = (
                    transformar_demanda_dtm(raw_path)
                    if raw_path and os.path.exists(raw_path)
                    else pd.DataFrame()
                )
            estoque                  = passo_3_estoque()
            entradas, df_abertos_fut = passo_4_pedidos_abertos()
            abc                      = passo_5_abc(demanda, materiais, contratos=contratos)
            df_mrp, df_ped           = passos_6_11_mrp(
                demanda, estoque, entradas, abc, materiais,
                lead_time_dias=LEAD_TIME_DIAS,
                lead_times_dict=lt_dict or None,
            )

            demanda_detail = _demanda_detail_raw if not _demanda_detail_raw.empty else None
            df_rateio = passo_12_rateio(df_ped, df_abertos_fut, demanda_detail=demanda_detail)
            # demanda_detail salvo para filtros de departamento/programa na Projeção de Estoque
            _demanda_detail_df = demanda_detail if demanda_detail is not None else pd.DataFrame()

            # ── Histórico MB51 (opcional) ─────────────────────────────────────
            mb51_path = achar_arquivo("historico_mb51.csv")
            df_mb51 = ler_historico_mb51(mb51_path) if mb51_path else pd.DataFrame()

            # ── Política de Pagamento (opcional) ─────────────────────────────
            politica_path = achar_arquivo("politica_pagamento.csv") or achar_arquivo("Politica_de_pagamento")
            politica_pag_carregada: dict[str, list[int]] = (
                ler_politica_pagamento(politica_path)
                if politica_path else {}
            )

            # Alertas
            alertas_rup, alertas_cont = _calcular_alertas(df_mrp, df_ped, contratos)

            # ── Cálculos Financeiros (independente de qual aba o usuário visita) ──
            _linhas_fin: list[pd.DataFrame] = []

            if not df_ped.empty:
                _tmp = _classificar_contratos_mrp(df_ped, contratos)
                print(f"  [FIN] Novos pedidos MRP      : R$ {_tmp['valor_total_pedido'].sum():,.2f}"
                      f" ({len(df_ped)} pedidos → {len(_tmp)} linhas após cobertura contratual)")
                _tmp["mes_pedido"]           = pd.to_datetime(_tmp["data_pedido"], format="%d/%m/%Y", errors="coerce").dt.to_period("M").astype(str)
                _tmp["mes_entrega"]          = _tmp["periodo_entrega"]
                _tmp["data_base_pagamento"]  = pd.to_datetime(_tmp["data_chegada"], format="%d/%m/%Y", errors="coerce")
                _tmp["documento_referencia"] = None
                _tmp["numero_pedido"]        = None
                _tmp["origem"]               = _tmp["_origem_mrp"]
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
                        _tmp2["mes_remessa"] + "-15", format="%Y-%m-%d", errors="coerce"
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
                # mes_pedido: mês da data real do documento MB51; fallback = mes_entrega
                if "data_doc" in _tmp3.columns:
                    _tmp3["mes_pedido"] = (
                        pd.to_datetime(_tmp3["data_doc"], errors="coerce")
                        .dt.to_period("M").astype(str)
                    )
                    _mask_mb51_nd = _tmp3["mes_pedido"].isna() | (_tmp3["mes_pedido"] == "NaT")
                    _tmp3.loc[_mask_mb51_nd, "mes_pedido"] = _tmp3.loc[_mask_mb51_nd, "mes_entrega"]
                else:
                    _tmp3["mes_pedido"] = _tmp3["mes_entrega"]
                # Usar data real do documento MB51 se disponível; fallback dia 15 do mês
                if "data_doc" in _tmp3.columns:
                    _tmp3["data_base_pagamento"] = pd.to_datetime(_tmp3["data_doc"], errors="coerce")
                    _mask_no_date = _tmp3["data_base_pagamento"].isna()
                    _tmp3.loc[_mask_no_date, "data_base_pagamento"] = pd.to_datetime(
                        _tmp3.loc[_mask_no_date, "mes_entrega"] + "-15",
                        format="%Y-%m-%d", errors="coerce",
                    )
                else:
                    _tmp3["data_base_pagamento"] = pd.to_datetime(
                        _tmp3["mes_entrega"] + "-15", format="%Y-%m-%d", errors="coerce"
                    )
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
                            "origem"              : _row["origem"],
                            "material"            : _row["material"],
                            "mes_emissao"         : _row.get("mes_pedido", "-"),
                            "mes_entrega"         : _row.get("mes_entrega", "-"),
                            "prazo_dias"          : _d,
                            "mes_pagamento"       : (_data_base + timedelta(days=_d)).strftime("%Y-%m"),
                            "valor_pedido_total"  : _row["valor_pedido"],   # valor integral do pedido (sem split)
                            "valor_parcela"       : _resto if _i == _n - 1 else _vbase,
                            "num_parcela"         : _i + 1,
                            "tot_parcelas"        : _n,
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

    # ── Métricas resumo (sempre visíveis no topo) ─────────────────────────────
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

    # ── Pré-computar df_mrp_val (usado no download Excel) ────────────────────
    abc = r["abc"]
    df_mrp_val = df_mrp.merge(
        abc[["material", "valor_unitario"]].drop_duplicates("material"),
        on="material", how="left",
    )
    df_mrp_val["valor_unitario"] = df_mrp_val["valor_unitario"].fillna(0)
    df_mrp_val["valor_pedido"] = df_mrp_val["pedido_gerado"] * df_mrp_val["valor_unitario"]
    df_mrp_val = df_mrp_val.drop(columns=["valor_unitario"])

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

    if _pagina == "📅 Projeção de Estoque":
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

    if _pagina == "💰 Financeiro":
        st.subheader("Financeiro")

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

            # ── Filtros adicionais: Material e Classe ABC ─────────────────────
            _mats_fin_disp = sorted(
                df_fin_bruto["material"].dropna().astype(str).unique().tolist()
                if not df_fin_bruto.empty and "material" in df_fin_bruto.columns
                else []
            )
            _abc_fin    = r.get("abc", pd.DataFrame())
            _classe_map_fin = (
                dict(zip(_abc_fin["material"].astype(str), _abc_fin["classe"]))
                if not _abc_fin.empty and "classe" in _abc_fin.columns else {}
            )
            _classes_fin_disp = sorted(
                _abc_fin["classe"].dropna().unique().tolist()
                if not _abc_fin.empty and "classe" in _abc_fin.columns
                else ["A", "B", "C"]
            )
            _ff4, _ff5 = st.columns(2)
            with _ff4:
                _sel_mats = st.multiselect(
                    "Material", _mats_fin_disp, default=[],
                    key="fin_mats", placeholder="Todos os materiais",
                )
            with _ff5:
                _sel_classes = st.multiselect(
                    "Classe ABC", _classes_fin_disp, default=[],
                    key="fin_classes", placeholder="Todas as classes",
                )

            # ── Enriquecer com classe para filtro ─────────────────────────────
            def _add_classe(df: pd.DataFrame) -> pd.DataFrame:
                if df.empty or not _classe_map_fin:
                    return df
                out = df.copy()
                out["_classe"] = out["material"].astype(str).map(_classe_map_fin)
                return out

            _fin_bruto_c   = _add_classe(df_fin_bruto)
            _fluxo_bruto_c = _add_classe(df_fluxo_bruto)

            # ── Filtrar brutos ────────────────────────────────────────────────
            def _filtrar_fin(df: pd.DataFrame, col_mes: str) -> pd.DataFrame:
                if df.empty:
                    return df
                out = df[df[col_mes].str[:4] == _sel_ano].copy() if col_mes in df.columns else df.copy()
                if _sel_depto != "Todos" and "departamento" in out.columns:
                    out = out[out["departamento"] == _sel_depto]
                if _sel_prog != "Todos" and "programa_orcamentario" in out.columns:
                    out = out[out["programa_orcamentario"] == _sel_prog]
                if _sel_mats and "material" in out.columns:
                    out = out[out["material"].astype(str).isin(_sel_mats)]
                if _sel_classes and "_classe" in out.columns:
                    out = out[out["_classe"].isin(_sel_classes)]
                return out

            _fin_f   = _filtrar_fin(_fin_bruto_c,   "mes_pedido")
            _fluxo_f = _filtrar_fin(_fluxo_bruto_c, "mes_pagamento")

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

            subtab_orc, subtab_cx, subtab_ent = st.tabs([
                "📊 Visão Orçamentária (Emissão)",
                "💸 Visão de Caixa (Desembolso Real)",
                "📦 Entradas Mensais",
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
                    # ── Alerta de cobertura contratual ────────────────────────
                    _sem_cont_cols_orc = [_c for _c in _vis_orc_f.columns if "Sem Contrato" in str(_c)]
                    _com_cont_cols_orc = [_c for _c in _vis_orc_f.columns if "Com Contrato" in str(_c)]
                    if _sem_cont_cols_orc or _com_cont_cols_orc:
                        _orc_nodrop = _vis_orc_f.drop("TOTAL GERAL", errors="ignore")
                        _val_sem = _orc_nodrop[_sem_cont_cols_orc].sum().sum() if _sem_cont_cols_orc else 0.0
                        _val_com = _orc_nodrop[_com_cont_cols_orc].sum().sum() if _com_cont_cols_orc else 0.0
                        _val_mrp_total = _val_sem + _val_com
                        _col_a, _col_b = st.columns(2)
                        _col_a.metric(
                            "✅ Com Contrato (pode emitir)",
                            _fmt_brl_contabil(_val_com),
                            f"{_val_com/_val_mrp_total*100:.0f}% dos pedidos MRP" if _val_mrp_total > 0 else "—",
                        )
                        _col_b.metric(
                            "⚠️ Sem Contrato (depende de Compras)",
                            _fmt_brl_contabil(_val_sem),
                            f"{_val_sem/_val_mrp_total*100:.0f}% dos pedidos MRP" if _val_mrp_total > 0 else "—",
                            delta_color="inverse",
                        )
                        if _val_sem > 0:
                            st.warning(
                                f"**{_fmt_brl_contabil(_val_sem)}** em pedidos futuros só poderão ser emitidos "
                                f"quando a equipe de **Compras** disponibilizar contratos vigentes."
                            )

                    def _merge_mrp_para_grafico(_pv: pd.DataFrame) -> pd.DataFrame:
                        """Funde colunas MRP Com/Sem Contrato em 'Novo Pedido (MRP)' para o gráfico."""
                        _pv2 = _pv.copy()
                        _mrp_split = [_c for _c in _pv2.columns if "Novo Pedido (MRP)" in str(_c) and _c != "Novo Pedido (MRP)"]
                        if _mrp_split:
                            _pv2["Novo Pedido (MRP)"] = _pv2.get("Novo Pedido (MRP)", 0) + _pv2[_mrp_split].sum(axis=1)
                            _pv2 = _pv2.drop(columns=_mrp_split)
                        return _pv2

                    _chart_financeiro(_merge_mrp_para_grafico(_vis_orc_f.drop("TOTAL GERAL", errors="ignore")), "Compromisso por Mês de Emissão")
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
                    # ── Alerta de cobertura contratual (caixa) ────────────────
                    _sem_cont_cols_cx = [_c for _c in _vis_cx_f.columns if "Sem Contrato" in str(_c)]
                    _com_cont_cols_cx = [_c for _c in _vis_cx_f.columns if "Com Contrato" in str(_c)]
                    if _sem_cont_cols_cx or _com_cont_cols_cx:
                        _cx_nodrop = _vis_cx_f.drop("TOTAL GERAL", errors="ignore")
                        _cx_val_sem = _cx_nodrop[_sem_cont_cols_cx].sum().sum() if _sem_cont_cols_cx else 0.0
                        _cx_val_com = _cx_nodrop[_com_cont_cols_cx].sum().sum() if _com_cont_cols_cx else 0.0
                        _cx_mrp_tot = _cx_val_sem + _cx_val_com
                        _cxc1, _cxc2 = st.columns(2)
                        _cxc1.metric(
                            "✅ Desembolso c/ Contrato",
                            _fmt_brl_contabil(_cx_val_com),
                            f"{_cx_val_com/_cx_mrp_tot*100:.0f}% dos pedidos MRP" if _cx_mrp_tot > 0 else "—",
                        )
                        _cxc2.metric(
                            "⚠️ Desembolso s/ Contrato",
                            _fmt_brl_contabil(_cx_val_sem),
                            f"{_cx_val_sem/_cx_mrp_tot*100:.0f}% dos pedidos MRP" if _cx_mrp_tot > 0 else "—",
                            delta_color="inverse",
                        )
                        if _cx_val_sem > 0:
                            st.warning(
                                f"**{_fmt_brl_contabil(_cx_val_sem)}** do desembolso previsto está **condicionado "
                                f"à contratação** pela equipe de Compras."
                            )

                    _chart_financeiro(_merge_mrp_para_grafico(_vis_cx_f.drop("TOTAL GERAL", errors="ignore")), "Desembolso por Mês de Pagamento")

                    # ── Diagnóstico automático de picos ───────────────────────
                    _cx_totais = (
                        _vis_cx_f.drop("TOTAL GERAL", errors="ignore")["Total"]
                        if "Total" in _vis_cx_f.columns else pd.Series(dtype=float)
                    )
                    if not _cx_totais.empty and len(_cx_totais) >= 2:
                        _media_cx = _cx_totais.mean()
                        _std_cx   = _cx_totais.std()
                        _picos_cx_idx = set(
                            _cx_totais[_cx_totais > _media_cx + _std_cx].index.tolist()
                        )
                        _n_picos = len(_picos_cx_idx)
                        _exp_title = (
                            f"📊 Resumo Mensal — {len(_cx_totais)} mês(es)"
                            + (f" · ⚠️ {_n_picos} pico(s) acima do normal" if _n_picos else "")
                        )
                        with st.expander(_exp_title, expanded=False):
                            _desc_pico_all = r.get("materiais_df", pd.DataFrame())
                            _dp_map_all = dict(zip(
                                _desc_pico_all["material"].astype(str),
                                _desc_pico_all["descricao"]
                            )) if not _desc_pico_all.empty and "descricao" in _desc_pico_all.columns else {}

                            for _p_mes, _p_val in _cx_totais.sort_index().items():
                                _is_pico = _p_mes in _picos_cx_idx
                                _fator = _p_val / _media_cx if _media_cx > 0 else 0
                                _pico_badge = " ⚠️ pico" if _is_pico else ""
                                st.markdown(
                                    f"#### {'📌' if _is_pico else '📅'} {_p_mes} — {_fmt_brl_contabil(_p_val)} "
                                    f"&nbsp;·&nbsp; {_fator:.1f}× a média{_pico_badge}"
                                )

                                if not _fluxo_f.empty and "mes_pagamento" in _fluxo_f.columns:
                                    _pd = _fluxo_f[_fluxo_f["mes_pagamento"] == _p_mes]

                                    # Composição por origem
                                    _orig_pico = (
                                        _pd.groupby("origem")["valor_rateado"]
                                        .sum().sort_values(ascending=False)
                                    )
                                    if not _orig_pico.empty:
                                        _comp_parts = [
                                            f"{_o}: **{_fmt_brl_contabil(_v)}** ({_v/_p_val*100:.0f}%)"
                                            for _o, _v in _orig_pico.items()
                                        ]
                                        st.markdown("**Composição:** " + " · ".join(_comp_parts))

                                    # Explicação causal (para MRP)
                                    _mrp_pd = _pd[_pd["origem"].str.contains("MRP", na=False)]
                                    if not _mrp_pd.empty:
                                        _ems = sorted(_mrp_pd["mes_emissao"].dropna().unique()) \
                                            if "mes_emissao" in _mrp_pd.columns else []
                                        _ents = sorted(_mrp_pd["mes_entrega"].dropna().unique()) \
                                            if "mes_entrega" in _mrp_pd.columns else []
                                        _prazos = sorted(
                                            _mrp_pd["prazo_dias"].dropna().astype(int).unique()
                                        ) if "prazo_dias" in _mrp_pd.columns else []
                                        _n_mat = _mrp_pd["material"].nunique() \
                                            if "material" in _mrp_pd.columns else 0

                                        st.info(
                                            f"**Por que {_p_mes}?**  \n"
                                            f"O sistema MRP gerou pedidos para **{_n_mat} material(is)** "
                                            f"nos meses **{', '.join(_ems) if _ems else '?'}** "
                                            f"antecipando o lead time.  \n"
                                            f"Esses pedidos chegam ao estoque em "
                                            f"**{', '.join(_ents) if _ents else '?'}** "
                                            f"e, com prazo(s) contratual(is) de "
                                            f"**{' / '.join(str(p)+'d' for p in _prazos) if _prazos else '60/90d'}**, "
                                            f"os pagamentos se concentram em **{_p_mes}**.  \n"
                                            + (
                                                f"ℹ️ O valor total anual **não muda** — é uma concentração "
                                                f"de timing, não um aumento de gasto."
                                                if _is_pico else ""
                                            )
                                        )

                                    # Top 5 materiais
                                    _mat_pico = (
                                        _pd.groupby("material")["valor_rateado"]
                                        .sum().sort_values(ascending=False).head(5)
                                    )
                                    if not _mat_pico.empty:
                                        st.markdown("**Top materiais:**")
                                        for _mp_mat, _mp_val in _mat_pico.items():
                                            _mp_desc = _dp_map_all.get(str(_mp_mat), "")
                                            _label = f"{_mp_mat}" + (f" — {_mp_desc}" if _mp_desc else "")
                                            st.markdown(
                                                f"&nbsp;&nbsp;• {_label}: "
                                                f"**{_fmt_brl_contabil(_mp_val)}** "
                                                f"({_mp_val/_p_val*100:.0f}%)"
                                            )

                                st.divider()

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
                                # Gera coluna de parcela legível antes de filtrar colunas
                                if "num_parcela" in _detail_cx.columns and "tot_parcelas" in _detail_cx.columns:
                                    _detail_cx["Parcela"] = (
                                        _detail_cx["num_parcela"].astype(int).astype(str)
                                        + "º de "
                                        + _detail_cx["tot_parcelas"].astype(int).astype(str)
                                    )
                                else:
                                    _detail_cx["Parcela"] = "-"

                                _cx_dcols = [c for c in [
                                    "material", "origem",
                                    "mes_emissao",        # data geração pedido
                                    "mes_entrega",        # data chegada do material
                                    "Parcela",            # 1º de 2, 2º de 2, ...
                                    "prazo_dias",         # prazo de pagamento
                                    "mes_pagamento",      # data desembolso
                                    "departamento", "programa_orcamentario",
                                    "valor_pedido_total", # valor integral do PO
                                    "valor_parcela",      # valor da parcela (antes rateio)
                                    "valor_rateado",      # valor da parcela após rateio
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
                                        "valor_pedido_total"    : "Valor Total Pedido",
                                        "valor_parcela"         : "Valor da Parcela",
                                        "valor_rateado"         : "Valor Parcela Rateado",
                                    })
                                    .sort_values(["Data Chegada", "Material"])
                                    .reset_index(drop=True)
                                )
                                _fmt_cx = {c: _fmt_brl_contabil for c in [
                                    "Valor Total Pedido", "Valor da Parcela", "Valor Parcela Rateado"
                                ] if c in _detail_cx_show.columns}
                                _tot_cx = _detail_cx_show["Valor Parcela Rateado"].sum() if "Valor Parcela Rateado" in _detail_cx_show.columns else 0.0
                                st.caption(
                                    f"{len(_detail_cx_show)} parcela(s) · "
                                    f"Total rateado: **{_fmt_brl_contabil(_tot_cx)}**"
                                )
                                st.dataframe(
                                    _detail_cx_show.style.format(_fmt_cx, na_rep="-"),
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
                            _fluxo_trace = _fluxo_f.copy()
                            if "num_parcela" in _fluxo_trace.columns and "tot_parcelas" in _fluxo_trace.columns:
                                _fluxo_trace["parcela_label"] = (
                                    _fluxo_trace["num_parcela"].astype(int).astype(str)
                                    + "º de "
                                    + _fluxo_trace["tot_parcelas"].astype(int).astype(str)
                                )
                            else:
                                _fluxo_trace["parcela_label"] = "-"

                            # Enriquecer com descrição do material
                            _mat_df_tr = r.get("materiais_df", pd.DataFrame())
                            _abc_tr    = r.get("abc", pd.DataFrame())
                            _desc_tr: dict = {}
                            if not _mat_df_tr.empty and "descricao" in _mat_df_tr.columns:
                                _desc_tr = dict(zip(_mat_df_tr["material"].astype(str), _mat_df_tr["descricao"]))
                            elif not _abc_tr.empty and "descricao" in _abc_tr.columns:
                                _desc_tr = dict(zip(_abc_tr["material"].astype(str), _abc_tr["descricao"]))
                            _fluxo_trace["descricao"] = (
                                _fluxo_trace["material"].astype(str).map(_desc_tr).fillna("-")
                            )

                            _trace_cols = [c for c in [
                                "origem", "material", "descricao", "departamento", "programa_orcamentario",
                                "mes_emissao", "mes_entrega",
                                "parcela_label", "prazo_dias", "mes_pagamento",
                                "valor_pedido_total", "valor_parcela", "valor_rateado",
                            ] if c in _fluxo_trace.columns]
                            df_trace = (
                                _fluxo_trace[_trace_cols]
                                .rename(columns={
                                    "origem"                : "Origem",
                                    "material"              : "Material",
                                    "descricao"             : "Descrição",
                                    "departamento"          : "Departamento",
                                    "programa_orcamentario" : "Programa",
                                    "mes_emissao"           : "Mês Emissão",
                                    "mes_entrega"           : "Mês Entrega",
                                    "parcela_label"         : "Parcela",
                                    "prazo_dias"            : "Prazo (dias)",
                                    "mes_pagamento"         : "Mês Pagamento",
                                    "valor_pedido_total"    : "Valor Total Pedido",
                                    "valor_parcela"         : "Valor da Parcela",
                                    "valor_rateado"         : "Valor Parcela Rateado",
                                })
                                .sort_values(["Mês Pagamento", "Mês Entrega"])
                                .reset_index(drop=True)
                            )
                            _fmt_trace = {c: _fmt_brl_contabil for c in [
                                "Valor Total Pedido", "Valor da Parcela", "Valor Parcela Rateado"
                            ] if c in df_trace.columns}
                            st.dataframe(
                                df_trace.style.format(_fmt_trace, na_rep="-"),
                                use_container_width=True,
                                height=400,
                            )

            with subtab_ent:
                st.caption(
                    f"Valor e quantidade de pedidos por mês de chegada ao estoque · {_sel_ano}"
                    + (f" · Departamento: {_sel_depto}" if _sel_depto != "Todos" else "")
                    + (f" · Programa: {_sel_prog}" if _sel_prog != "Todos" else "")
                )
                _ent_f = _filtrar_fin(_fin_bruto_c, "mes_entrega")
                if _ent_f.empty:
                    st.info("Sem dados de entradas para os filtros selecionados.")
                else:
                    _ent_grp = (
                        _ent_f
                        .groupby("mes_entrega", as_index=False)
                        .agg(
                            valor_total  = ("valor_rateado", "sum"),
                            qtd_pedidos  = ("material",      "count"),
                            qtd_materiais= ("material",      "nunique"),
                        )
                        .sort_values("mes_entrega")
                    )
                    # Gráfico
                    _fig_ent = go.Figure(go.Bar(
                        x=_ent_grp["mes_entrega"],
                        y=_ent_grp["valor_total"],
                        text=_ent_grp["valor_total"].apply(
                            lambda v: f"R$ {v/1_000:.1f}K" if v < 1_000_000
                            else f"R$ {v/1_000_000:.2f}M"
                        ),
                        textposition="outside",
                        marker_color="#1f77b4",
                    ))
                    _fig_ent.update_layout(
                        title="Entradas Mensais — Valor Total (R$)",
                        yaxis_title="R$",
                        height=360,
                        margin=dict(t=50, b=40),
                    )
                    st.plotly_chart(_fig_ent, use_container_width=True)
                    # Tabela
                    _ent_tbl = _ent_grp.copy()
                    _ent_tbl["Valor Total"] = _ent_tbl["valor_total"].apply(_fmt_brl_contabil)
                    st.dataframe(
                        _ent_tbl.rename(columns={
                            "mes_entrega"   : "Mês",
                            "Valor Total"   : "Valor Total (R$)",
                            "qtd_pedidos"   : "Qtd. Linhas",
                            "qtd_materiais" : "Materiais Distintos",
                        })[["Mês", "Valor Total (R$)", "Qtd. Linhas", "Materiais Distintos"]],
                        use_container_width=True,
                        hide_index=True,
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

    if _pagina == "📋 Saldo de Contrato":
        st.subheader("Cobertura Contratual dos Pedidos MRP")
        if contratos.empty:
            st.info(
                "Arquivo de contratos SAP não carregado.  \n"
                "Carregue o arquivo **Contratos_SAP** (ME3M/ME3N, separado por TAB) para "
                "ver quais pedidos estão cobertos e quais dependem de novos contratos."
            )
        else:
            # ── Monta tabela de cobertura por material ────────────────────────
            _fin_b = r.get("df_fin_bruto", pd.DataFrame())
            _mrp_b = (
                _fin_b[_fin_b["origem"].str.contains("MRP", na=False)].copy()
                if not _fin_b.empty and "origem" in _fin_b.columns else pd.DataFrame()
            )

            if not _mrp_b.empty:
                _mrp_b["_coberto"] = _mrp_b["origem"].str.contains("Com Contrato", na=False)
                _cob = (
                    _mrp_b.groupby("material").apply(
                        lambda _g: pd.Series({
                            "Pedidos c/ Contrato (R$)": _g.loc[_g["_coberto"], "valor_pedido"].sum(),
                            "Pedidos s/ Contrato (R$)": _g.loc[~_g["_coberto"], "valor_pedido"].sum(),
                            "Qtd c/ Contrato"         : _g.loc[_g["_coberto"], "quantidade"].sum(),
                            "Qtd s/ Contrato"         : _g.loc[~_g["_coberto"], "quantidade"].sum(),
                        })
                    ).reset_index()
                )
                # Enriquecer com dados do arquivo de contratos
                _cont_cols = ["material", "tipo_contrato", "saldo_contrato", "data_fim_vigencia"]
                _cont_cols = [_c for _c in _cont_cols if _c in contratos.columns]
                _cob = _cob.merge(contratos[_cont_cols], on="material", how="left")

                # Observação: ACORDO DE PREÇO (1 contrato por material) ou ACORDO DE QUANTIDADE (>1)
                _n_cont_por_mat = contratos.groupby("material").size().reset_index(name="_n_contratos")
                _cob = _cob.merge(_n_cont_por_mat, on="material", how="left")
                _cob["Observação"] = _cob["_n_contratos"].apply(
                    lambda n: "ACORDO DE PREÇO" if n == 1 else ("ACORDO DE QUANTIDADE" if n > 1 else "")
                )
                _cob = _cob.drop(columns=["_n_contratos"])

                # Enriquecer com descrição se disponível
                _mat_df_c = r.get("materiais_df", pd.DataFrame())
                if not _mat_df_c.empty and "descricao" in _mat_df_c.columns:
                    _desc_c = _mat_df_c[["material","descricao"]].drop_duplicates("material")
                    _cob = _cob.merge(_desc_c, on="material", how="left")
                    _cob.insert(1, "Descrição", _cob.pop("descricao"))

                _cob["Total MRP (R$)"] = _cob["Pedidos c/ Contrato (R$)"] + _cob["Pedidos s/ Contrato (R$)"]
                _cob["Cobertura %"]    = (
                    (_cob["Pedidos c/ Contrato (R$)"] / _cob["Total MRP (R$)"] * 100)
                    .where(_cob["Total MRP (R$)"] > 0, 0)
                    .round(1)
                )

                # Totais gerais
                _tot_com = _cob["Pedidos c/ Contrato (R$)"].sum()
                _tot_sem = _cob["Pedidos s/ Contrato (R$)"].sum()
                _tot_mrp = _tot_com + _tot_sem
                _n_sem   = (_cob["Pedidos s/ Contrato (R$)"] > 0).sum()

                if _tot_sem > 0:
                    st.warning(
                        f"**{_n_sem} material(is)** sem cobertura contratual total ou parcial — "
                        f"**{_fmt_brl_contabil(_tot_sem)}** dependem de novos contratos de Compras "
                        f"({_tot_sem/_tot_mrp*100:.0f}% do total MRP)."
                    )
                else:
                    st.success("✅ 100% dos pedidos MRP estão cobertos por contratos vigentes.")

                _mc1, _mc2, _mc3 = st.columns(3)
                _mc1.metric("Total MRP", _fmt_brl_contabil(_tot_mrp))
                _mc2.metric("✅ Com Contrato", _fmt_brl_contabil(_tot_com))
                _mc3.metric("⚠️ Sem Contrato", _fmt_brl_contabil(_tot_sem))

                st.markdown("#### Detalhamento por Material")

                # Ordenar: sem contrato primeiro
                _cob = _cob.sort_values("Pedidos s/ Contrato (R$)", ascending=False).reset_index(drop=True)

                # Máscara numérica antes de formatar
                _sem_mask = (_cob["Pedidos s/ Contrato (R$)"] > 0).to_dict()

                def _style_cont_row(_row):
                    if _sem_mask.get(_row.name, False):
                        return ["background-color: #fff3cd"] * len(_row)
                    return [""] * len(_row)

                _cob_fmt = _cob.copy()
                for _fc in ["Pedidos c/ Contrato (R$)", "Pedidos s/ Contrato (R$)", "Total MRP (R$)"]:
                    if _fc in _cob_fmt.columns:
                        _cob_fmt[_fc] = _cob_fmt[_fc].apply(_fmt_brl_contabil)
                for _fc in ["Qtd c/ Contrato", "Qtd s/ Contrato"]:
                    if _fc in _cob_fmt.columns:
                        _cob_fmt[_fc] = _cob_fmt[_fc].apply(lambda _v: f"{_v:,.0f}".replace(",","."))
                if "Cobertura %" in _cob_fmt.columns:
                    _cob_fmt["Cobertura %"] = _cob_fmt["Cobertura %"].apply(lambda _v: f"{_v:.1f}%")

                st.dataframe(
                    _cob_fmt.style.apply(_style_cont_row, axis=1),
                    use_container_width=True,
                    height=420,
                )

                st.caption(
                    "🟡 Linhas destacadas = material com pedidos sem cobertura contratual.  \n"
                    "**Acordo de Preço**: qtd ilimitada, verificar vigência.  \n"
                    "**Acordo de Quantidade**: limitado ao saldo disponível."
                )

            # ── Contratos vigentes carregados ─────────────────────────────────
            with st.expander("📄 Contratos vigentes carregados (raw)", expanded=False):
                st.dataframe(contratos, use_container_width=True)

    if _pagina == "📂 Rateio":
        st.subheader("Rateio por Departamento / Programa Orçamentário")
        if df_rateio.empty:
            st.info("Nenhum dado de rateio disponível.")
        else:
            st.dataframe(df_rateio, use_container_width=True, height=380)

    if _pagina == "⚠️ Rateio Pendente":
        st.subheader("Rateio Pendente / Atribuição Manual")

        # ── Auxiliares ────────────────────────────────────────────────────────
        _mat_df   = r.get("materiais_df", pd.DataFrame())
        _abc_df   = r.get("abc",          pd.DataFrame())
        _dem_df   = r.get("demanda_df",   pd.DataFrame())
        _dd_df    = r.get("demanda_detail_df", pd.DataFrame())

        _desc_map : dict = {}
        if not _mat_df.empty and "descricao" in _mat_df.columns:
            _desc_map = dict(zip(_mat_df["material"].astype(str), _mat_df["descricao"]))

        _preco_map: dict = {}
        if not _abc_df.empty and "valor_unitario" in _abc_df.columns:
            _preco_map = dict(zip(_abc_df["material"].astype(str), _abc_df["valor_unitario"]))

        # ── Seção 1: Log de materiais sem rateio ─────────────────────────────
        st.markdown("#### 📋 Materiais sem rateio definido")

        if df_rateio.empty or "departamento" not in df_rateio.columns:
            st.info("Nenhum dado de rateio disponível. Processe o MRP primeiro.")
        else:
            _df_nao_def = df_rateio[df_rateio["departamento"] == "NAO_DEFINIDO"].copy()
            _mats_nao_def = sorted(_df_nao_def["material"].astype(str).unique().tolist()) \
                if not _df_nao_def.empty else []

            n_pend = len(_mats_nao_def)
            if n_pend > 0:
                st.warning(f"⚠️ **{n_pend} material(is) pendente(s) de rateio manual**")
            else:
                st.success("✅ Todos os materiais têm rateio definido.")

            if not _df_nao_def.empty:
                def _motivo(mat: str) -> str:
                    mat = str(mat)
                    if not _dem_df.empty and mat not in _dem_df["material"].astype(str).values:
                        return "sem demanda cadastrada"
                    if not _dd_df.empty and "departamento" in _dd_df.columns:
                        _m = _dd_df[_dd_df["material"].astype(str) == mat]
                        if _m.empty or _m["departamento"].fillna("").str.strip().eq("").all():
                            return "demanda sem departamento"
                    return "material não mapeado"

                _sumario = (
                    _df_nao_def.groupby("material", as_index=False)
                    .agg(qtd_total_rateada=("qtd_rateada", "sum"))
                )
                _sumario["material"]    = _sumario["material"].astype(str)
                _sumario["descricao"]   = _sumario["material"].map(_desc_map).fillna("-")
                _sumario["valor_unit"]  = _sumario["material"].map(_preco_map).fillna(0.0)
                _sumario["valor_total"] = (_sumario["qtd_total_rateada"] * _sumario["valor_unit"]).round(2)
                _sumario["motivo"]      = _sumario["material"].apply(_motivo)
                st.dataframe(
                    _sumario[["material", "descricao", "qtd_total_rateada", "valor_total", "motivo"]],
                    use_container_width=True,
                    hide_index=True,
                )

        st.divider()

        # ── Seção 2: Formulário de atribuição manual ──────────────────────────
        st.markdown("#### ✏️ Atribuição Manual por Material")

        _rm_df = _carregar_rateio_manual()
        _mats_manual = sorted(
            _rm_df["material"].astype(str).unique().tolist()
        ) if not _rm_df.empty and "material" in _rm_df.columns else []

        _mats_opcoes = sorted(set(
            (locals().get("_mats_nao_def") or []) + _mats_manual
        ))

        if not _mats_opcoes:
            st.info("Nenhum material disponível para rateio manual.")
        else:
            mat_sel = st.selectbox(
                "Selecionar material",
                _mats_opcoes,
                key="rm_mat_sel",
                help="Materiais com NAO_DEFINIDO + materiais já com rateio manual (para edição)",
            )

            if mat_sel:
                _desc_sel = _desc_map.get(str(mat_sel), "-")
                _qtd_nao_def = 0.0
                if "locals" in dir() and not df_rateio.empty:
                    _nao_def_sub = df_rateio[
                        (df_rateio["material"].astype(str) == str(mat_sel)) &
                        (df_rateio["departamento"] == "NAO_DEFINIDO")
                    ]
                    _qtd_nao_def = _nao_def_sub["qtd_rateada"].sum()
                st.caption(f"Descrição: **{_desc_sel}** | Qtd pendente: **{_qtd_nao_def:,.2f} un.**")

                # Listas de departamentos/programas disponíveis
                _depts_set: set[str] = set()
                _progs_set: set[str] = set()
                for _src_df in (df_rateio, _dd_df):
                    if not _src_df.empty:
                        if "departamento" in _src_df.columns:
                            _depts_set |= set(
                                _src_df["departamento"]
                                .dropna().astype(str).str.strip()
                                .unique().tolist()
                            )
                        if "programa_orcamentario" in _src_df.columns:
                            _progs_set |= set(
                                _src_df["programa_orcamentario"]
                                .dropna().astype(str).str.strip()
                                .unique().tolist()
                            )
                _depts_list = sorted(
                    d for d in _depts_set if d not in ("NAO_DEFINIDO", "")
                ) or ["DPC"]
                _progs_list = sorted(
                    p for p in _progs_set if p not in ("NAO_DEFINIDO", "")
                ) or ["INDEFINIDO"]
                _depts_opts = _depts_list + ["Outro"]
                _progs_opts = _progs_list + ["Outro"]

                # Valores pré-preenchidos do rateio manual existente
                _existing_rows: list[dict] = []
                if not _rm_df.empty and "material" in _rm_df.columns:
                    _ex = _rm_df[_rm_df["material"].astype(str) == str(mat_sel)]
                    _existing_rows = _ex.to_dict("records")

                # Session state: número de linhas do formulário
                _linhas_key = f"rm_n_linhas_{mat_sel}"
                if _linhas_key not in st.session_state:
                    st.session_state[_linhas_key] = max(1, len(_existing_rows))

                n_linhas = st.session_state[_linhas_key]
                linhas_form: list[dict] = []

                for _i in range(n_linhas):
                    _c1, _c2, _c3 = st.columns([3, 3, 1])
                    _ex_row = _existing_rows[_i] if _i < len(_existing_rows) else {}

                    # Departamento
                    _def_dept = str(_ex_row.get("departamento", "")) if _ex_row else ""
                    _dept_idx = _depts_opts.index(_def_dept) if _def_dept in _depts_opts else 0
                    _dept_sel = _c1.selectbox(
                        f"Departamento {_i+1}", _depts_opts,
                        index=_dept_idx, key=f"rm_dept_{mat_sel}_{_i}",
                    )
                    if _dept_sel == "Outro":
                        _dept_val = _c1.text_input(
                            f"Departamento customizado {_i+1}",
                            key=f"rm_dept_outro_{mat_sel}_{_i}",
                        )
                    else:
                        _dept_val = _dept_sel

                    # Programa orçamentário
                    _def_prog = str(_ex_row.get("programa_orcamentario", "")) if _ex_row else ""
                    _prog_idx = _progs_opts.index(_def_prog) if _def_prog in _progs_opts else 0
                    _prog_sel = _c2.selectbox(
                        f"Programa {_i+1}", _progs_opts,
                        index=_prog_idx, key=f"rm_prog_{mat_sel}_{_i}",
                    )
                    if _prog_sel == "Outro":
                        _prog_val = _c2.text_input(
                            f"Programa customizado {_i+1}",
                            key=f"rm_prog_outro_{mat_sel}_{_i}",
                        )
                    else:
                        _prog_val = _prog_sel

                    # Percentual
                    _def_pct  = int(round(float(_ex_row.get("proporcao", 0)) * 100)) if _ex_row else 0
                    _pct_val  = _c3.number_input(
                        f"% {_i+1}", min_value=0, max_value=100,
                        step=5, value=_def_pct,
                        key=f"rm_pct_{mat_sel}_{_i}",
                    )

                    linhas_form.append({
                        "departamento"          : _dept_val,
                        "programa_orcamentario" : _prog_val,
                        "proporcao_pct"         : _pct_val,
                    })

                # Botão para adicionar linha
                if st.button("➕ Adicionar linha", key=f"rm_add_{mat_sel}"):
                    st.session_state[_linhas_key] = n_linhas + 1
                    st.rerun()

                # Validação e ações
                _soma = sum(l["proporcao_pct"] for l in linhas_form)
                if _soma != 100:
                    st.error(f"❌ Soma dos percentuais = **{_soma}%** (deve ser exatamente 100%)")

                _col_s, _col_r = st.columns(2)
                if _col_s.button(
                    "💾 Salvar",
                    key=f"rm_salvar_{mat_sel}",
                    disabled=(_soma != 100),
                    use_container_width=True,
                ):
                    _salvar_rateio_manual(str(mat_sel), linhas_form)
                    st.success(f"✅ Rateio manual salvo para **{mat_sel}**. Reprocessando MRP...")
                    st.session_state.pop("resultado", None)
                    st.rerun()

                if _col_r.button(
                    "🗑 Remover rateio manual",
                    key=f"rm_remover_{mat_sel}",
                    use_container_width=True,
                ):
                    _df_rm_del = _carregar_rateio_manual()
                    _df_rm_del = _df_rm_del[_df_rm_del["material"].astype(str) != str(mat_sel)]
                    os.makedirs(DIR_DADOS, exist_ok=True)
                    _df_rm_del.to_csv(_RATEIO_MANUAL_PATH, sep=";", index=False)
                    if _linhas_key in st.session_state:
                        del st.session_state[_linhas_key]
                    st.success(f"✅ Rateio manual removido para **{mat_sel}**. Reprocessando MRP...")
                    st.session_state.pop("resultado", None)
                    st.rerun()

        st.divider()

        # ── Seção 3: Importação em lote ───────────────────────────────────────
        st.markdown("#### 📤 Importação em Lote")
        st.caption(
            "Formato CSV esperado: `material;departamento;programa_orcamentario;proporcao`  \n"
            "Coluna `proporcao` aceita 0-1 (ex.: 0.70) ou 0-100 (ex.: 70).  \n"
            "A soma por material deve ser exatamente 100%."
        )
        f_rateio_lote = st.file_uploader(
            "Upload CSV de rateio em lote",
            type=["csv", "txt"],
            key="up_rateio_lote",
        )
        if f_rateio_lote:
            try:
                f_rateio_lote.seek(0)
                df_lote = pd.read_csv(f_rateio_lote, sep=";", dtype={"material": str})
                df_lote.columns = [c.strip().lower() for c in df_lote.columns]
                # Normalizar nomes de coluna (case/espaço)
                _col_rename = {
                    "mat"      : "material",
                    "dept"     : "departamento",
                    "prog"     : "programa_orcamentario",
                    "prop"     : "proporcao",
                    "percent"  : "proporcao",
                    "pct"      : "proporcao",
                    "percentual": "proporcao",
                }
                df_lote.rename(columns=_col_rename, inplace=True)
                if "material" not in df_lote.columns:
                    expected = ["material", "departamento", "programa_orcamentario", "proporcao"]
                    df_lote.columns = expected[:len(df_lote.columns)]
                df_lote["material"] = df_lote["material"].astype(str).str.strip()
                df_lote["proporcao"] = pd.to_numeric(df_lote["proporcao"], errors="coerce").fillna(0.0)
                # Detectar escala (0-1 vs 0-100)
                if df_lote["proporcao"].max() <= 1.0:
                    df_lote["proporcao_pct"] = (df_lote["proporcao"] * 100).round(4)
                else:
                    df_lote["proporcao_pct"] = df_lote["proporcao"].round(4)
                # Validar soma por material
                _erros_lote: list[str] = []
                for _mat_l, _grp_l in df_lote.groupby("material"):
                    _soma_l = round(_grp_l["proporcao_pct"].sum(), 2)
                    if abs(_soma_l - 100.0) > 0.1:
                        _erros_lote.append(f"**{_mat_l}**: soma = {_soma_l}% ≠ 100%")
                if _erros_lote:
                    for _e in _erros_lote:
                        st.error(f"❌ {_e}")
                else:
                    st.success(f"✅ {df_lote['material'].nunique()} material(is) válido(s) — prévia:")
                    st.dataframe(
                        df_lote[["material", "departamento", "programa_orcamentario", "proporcao_pct"]],
                        use_container_width=True,
                        hide_index=True,
                    )
                    if st.button("✅ Confirmar importação em lote", key="rm_confirmar_lote"):
                        for _mat_l in df_lote["material"].unique():
                            _grp_l = df_lote[df_lote["material"] == _mat_l]
                            _lns = _grp_l[["departamento", "programa_orcamentario", "proporcao_pct"]].to_dict("records")
                            _salvar_rateio_manual(_mat_l, _lns)
                        st.success(f"✅ {df_lote['material'].nunique()} material(is) importado(s). Reprocessando MRP...")
                        st.session_state.pop("resultado", None)
                        st.rerun()
            except Exception as _e_lote:
                st.error(f"Erro ao ler arquivo: {_e_lote}")

    if _pagina == "📊 Rateio DFP - Realizado":
        st.subheader("Rateio DFP — Realizado")
        st.caption("Alocação Financeira por Material e Departamento")

        _dfp_rateio = r.get("df_rateio", pd.DataFrame()) if not df_rateio.empty else df_rateio

        if _dfp_rateio.empty:
            st.info(
                "Nenhum dado de alocação financeira disponível.  \n"
                "Processe o MRP com os arquivos de rateio carregados para visualizar a alocação realizada."
            )
        else:
            # Agrupar por material e departamento
            _dfp_cols_grp = [c for c in ["material", "departamento", "programa_orcamentario"] if c in _dfp_rateio.columns]
            _dfp_cols_num = [c for c in ["qtd_rateada", "valor_rateado"] if c in _dfp_rateio.columns]

            if _dfp_cols_grp and _dfp_cols_num:
                _dfp_agg = (
                    _dfp_rateio
                    .groupby(_dfp_cols_grp, as_index=False)[_dfp_cols_num]
                    .sum()
                    .sort_values(_dfp_cols_grp)
                    .reset_index(drop=True)
                )

                # Enriquecer com descrição do material
                _mat_dfp = r.get("materiais_df", pd.DataFrame())
                if not _mat_dfp.empty and "descricao" in _mat_dfp.columns:
                    _desc_dfp = _mat_dfp[["material", "descricao"]].drop_duplicates("material")
                    _dfp_agg = _dfp_agg.merge(_desc_dfp, on="material", how="left")
                    _dfp_agg.insert(1, "Descrição", _dfp_agg.pop("descricao"))

                # Filtros
                _dfp_depts = sorted(_dfp_agg["departamento"].dropna().unique().tolist()) if "departamento" in _dfp_agg.columns else []
                _dfp_f1, _dfp_f2 = st.columns(2)
                _dfp_sel_dept = _dfp_f1.selectbox("Departamento", ["Todos"] + _dfp_depts, key="dfp_dept")
                _dfp_mats = sorted(_dfp_agg["material"].dropna().astype(str).unique().tolist())
                _dfp_sel_mats = _dfp_f2.multiselect("Material", _dfp_mats, default=[], key="dfp_mats", placeholder="Todos")

                _dfp_show = _dfp_agg.copy()
                if _dfp_sel_dept != "Todos":
                    _dfp_show = _dfp_show[_dfp_show["departamento"] == _dfp_sel_dept]
                if _dfp_sel_mats:
                    _dfp_show = _dfp_show[_dfp_show["material"].astype(str).isin(_dfp_sel_mats)]

                # Formatar valores
                _dfp_fmt = _dfp_show.copy()
                if "valor_rateado" in _dfp_fmt.columns:
                    _dfp_fmt["valor_rateado"] = _dfp_fmt["valor_rateado"].apply(_fmt_brl_contabil)
                if "qtd_rateada" in _dfp_fmt.columns:
                    _dfp_fmt["qtd_rateada"] = _dfp_fmt["qtd_rateada"].apply(lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

                st.dataframe(_dfp_fmt, use_container_width=True, height=420, hide_index=True)

                # Totais
                _dfp_tot_val = _dfp_show["valor_rateado"].sum() if "valor_rateado" in _dfp_show.columns else 0.0
                _dfp_tot_mat = _dfp_show["material"].nunique()
                _dfp_tot_dept = _dfp_show["departamento"].nunique() if "departamento" in _dfp_show.columns else 0
                _dc1, _dc2, _dc3 = st.columns(3)
                _dc1.metric("Total Alocado", _fmt_brl_contabil(_dfp_tot_val))
                _dc2.metric("Materiais", _dfp_tot_mat)
                _dc3.metric("Departamentos", _dfp_tot_dept)
            else:
                st.dataframe(_dfp_rateio, use_container_width=True, height=380)

    if _pagina == "🎯 PCM — Materiais Críticos":
        st.subheader("PCM — Materiais Críticos")

        # Carregar arquivo PCM se disponível
        _pcm_path = _path_arquivo("pcm")
        _pcm_df = pd.DataFrame()
        if os.path.exists(_pcm_path):
            try:
                _pcm_df = pd.read_excel(_pcm_path)
                _pcm_df.columns = [str(c).strip() for c in _pcm_df.columns]
            except Exception as _e_pcm:
                st.warning(f"Erro ao carregar arquivo PCM: {_e_pcm}")

        _pcm_tabs = st.tabs([
            "📋 Agenda do Comprador",
            "🛒 Emitir Pedido de Compra",
            "📝 Contratar",
            "📊 Coberturas",
            "⚠️ Riscos por Dimensão",
            "👥 Atividades por Colaborador",
            "📤 Exportar Relatório",
        ])

        with _pcm_tabs[0]:  # Agenda do Comprador
            st.markdown("#### 📋 Agenda do Comprador")
            if _pcm_df.empty:
                st.info(
                    "Carregue o arquivo **PCM_*.xlsx** no painel lateral para visualizar "
                    "a agenda do comprador com os materiais críticos."
                )
            else:
                # Filtros básicos
                _ag_cols = _pcm_df.columns.tolist()
                _ag_f1, _ag_f2 = st.columns(2)
                _ag_mat_col = next((c for c in _ag_cols if "material" in c.lower() or "código" in c.lower()), None)
                _ag_resp_col = next((c for c in _ag_cols if "responsável" in c.lower() or "comprador" in c.lower() or "colaborador" in c.lower()), None)

                _ag_filter_mat = _ag_f1.text_input("Filtrar por material/código", key="ag_mat_filt")
                _ag_filter_resp = _ag_f2.text_input("Filtrar por comprador/responsável", key="ag_resp_filt")

                _ag_show = _pcm_df.copy()
                if _ag_filter_mat and _ag_mat_col:
                    _ag_show = _ag_show[_ag_show[_ag_mat_col].astype(str).str.contains(_ag_filter_mat, case=False, na=False)]
                if _ag_filter_resp and _ag_resp_col:
                    _ag_show = _ag_show[_ag_show[_ag_resp_col].astype(str).str.contains(_ag_filter_resp, case=False, na=False)]

                st.dataframe(_ag_show, use_container_width=True, height=420)
                st.caption(f"Total: {len(_ag_show)} registro(s)")

        with _pcm_tabs[1]:  # Emitir Pedido de Compra
            st.markdown("#### 🛒 Emitir Pedido de Compra")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para ver os materiais com pedido a emitir.")
            else:
                _emit_col = next((c for c in _pcm_df.columns if "emit" in c.lower() or "pedido" in c.lower() or "ação" in c.lower()), None)
                _emit_df = _pcm_df.copy()
                if _emit_col:
                    _emit_df = _emit_df[_emit_df[_emit_col].astype(str).str.contains("emitir|pedido", case=False, na=False)]
                st.dataframe(_emit_df, use_container_width=True, height=400)

        with _pcm_tabs[2]:  # Contratar
            st.markdown("#### 📝 Contratar")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para ver os materiais que precisam de contratação.")
            else:
                _contr_col = next((c for c in _pcm_df.columns if "contrat" in c.lower() or "ação" in c.lower()), None)
                _contr_df = _pcm_df.copy()
                if _contr_col:
                    _contr_df = _contr_df[_contr_df[_contr_col].astype(str).str.contains("contrat", case=False, na=False)]
                st.dataframe(_contr_df, use_container_width=True, height=400)

        with _pcm_tabs[3]:  # Coberturas
            st.markdown("#### 📊 Coberturas")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para visualizar as coberturas dos materiais críticos.")
            else:
                _cob_col = next((c for c in _pcm_df.columns if "cobertura" in c.lower() or "dias" in c.lower() or "meses" in c.lower()), None)
                if _cob_col:
                    _cob_pcm = _pcm_df[[c for c in _pcm_df.columns if True]].copy()
                    st.dataframe(_cob_pcm.sort_values(_cob_col) if _cob_col in _cob_pcm.columns else _cob_pcm,
                                 use_container_width=True, height=400)
                else:
                    st.dataframe(_pcm_df, use_container_width=True, height=400)

        with _pcm_tabs[4]:  # Riscos por Dimensão
            st.markdown("#### ⚠️ Riscos por Dimensão")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para visualizar os riscos por dimensão.")
            else:
                _risco_col = next((c for c in _pcm_df.columns if "risco" in c.lower() or "dimensão" in c.lower() or "criticidade" in c.lower()), None)
                if _risco_col:
                    _risco_grp = _pcm_df.groupby(_risco_col).size().reset_index(name="Qtd. Materiais")
                    _rc1, _rc2 = st.columns([1, 2])
                    _rc1.dataframe(_risco_grp, use_container_width=True, hide_index=True)
                    _rc2.dataframe(_pcm_df, use_container_width=True, height=350)
                else:
                    st.dataframe(_pcm_df, use_container_width=True, height=400)

        with _pcm_tabs[5]:  # Atividades por Colaborador
            st.markdown("#### 👥 Atividades por Colaborador")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para ver as atividades por colaborador.")
            else:
                _colab_col = next((c for c in _pcm_df.columns if "colaborador" in c.lower() or "comprador" in c.lower() or "responsável" in c.lower()), None)
                if _colab_col:
                    _colab_grp = _pcm_df.groupby(_colab_col).size().reset_index(name="Qtd. Materiais")
                    st.dataframe(_colab_grp, use_container_width=True, hide_index=True)
                    st.divider()
                    _sel_colab = st.selectbox("Ver atividades de:", ["Todos"] + sorted(_pcm_df[_colab_col].dropna().unique().tolist()), key="pcm_colab_sel")
                    _colab_show = _pcm_df if _sel_colab == "Todos" else _pcm_df[_pcm_df[_colab_col] == _sel_colab]
                    st.dataframe(_colab_show, use_container_width=True, height=360)
                else:
                    st.dataframe(_pcm_df, use_container_width=True, height=400)

        with _pcm_tabs[6]:  # Exportar Relatório
            st.markdown("#### 📤 Exportar Relatório PCM")
            if _pcm_df.empty:
                st.info("Carregue o arquivo **PCM_*.xlsx** para exportar o relatório.")
            else:
                try:
                    _pcm_excel = _gerar_excel({"PCM — Materiais Críticos": _pcm_df})
                    st.download_button(
                        label="⬇ Baixar Relatório PCM (Excel)",
                        data=_pcm_excel,
                        file_name=f"pcm_materiais_criticos_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as _e_pcm_exp:
                    st.error(f"Erro ao gerar relatório: {_e_pcm_exp}")

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
