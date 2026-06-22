import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="GFRT · FX Exposure",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── TEMA ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Fundo geral */
.stApp { background-color: #F7F6F3; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #1A2332;
    border-right: none;
}
[data-testid="stSidebar"] * { color: #C8D0DC !important; }
[data-testid="stSidebar"] .sidebar-logo {
    font-size: 22px; font-weight: 600;
    color: #FFFFFF !important;
    letter-spacing: -0.3px;
    padding: 8px 0 4px;
}
[data-testid="stSidebar"] .sidebar-sub {
    font-size: 11px; color: #6B7A8D !important;
    text-transform: uppercase; letter-spacing: .08em;
    margin-bottom: 20px;
}
[data-testid="stSidebar"] hr {
    border-color: #2C3A4D !important; margin: 12px 0;
}
[data-testid="stSidebar"] label { color: #8B99AA !important; font-size: 11px !important; }

/* Metric cards */
.kpi-card {
    background: #FFFFFF;
    border: 1px solid #ECEAE5;
    border-radius: 10px;
    padding: 18px 20px 14px;
}
.kpi-label { font-size: 11px; color: #8A8880; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 6px; }
.kpi-value { font-size: 26px; font-weight: 600; line-height: 1.1; color: #1A2332; }
.kpi-sub   { font-size: 11px; color: #A0A09A; margin-top: 5px; }
.kpi-pos   { color: #1A7A4A; }
.kpi-neg   { color: #B03A2E; }

/* Badge */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 11px; font-weight: 600;
    letter-spacing: .04em;
}
.badge-long  { background: #D6F0E3; color: #1A7A4A; }
.badge-short { background: #FAE0DC; color: #B03A2E; }

/* Page header */
.page-header {
    border-left: 3px solid #2E6B4F;
    padding-left: 14px;
    margin-bottom: 24px;
}
.page-header h2 { font-size: 20px; font-weight: 600; color: #1A2332; margin: 0 0 2px; }
.page-header p  { font-size: 12px; color: #8A8880; margin: 0; }

/* Section divider */
.sect { font-size: 11px; font-weight: 600; color: #8A8880;
        text-transform: uppercase; letter-spacing: .08em;
        border-bottom: 1px solid #ECEAE5; padding-bottom: 6px; margin: 24px 0 14px; }

/* Table styling override */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

COLORS = {
    "verde":   "#2E6B4F",
    "verde2":  "#1A7A4A",
    "azul":    "#1B4F8A",
    "laranja": "#C45C1A",
    "cinza":   "#6B7A8D",
    "amarelo": "#B8860B",
    "vermelho":"#B03A2E",
    "roxo":    "#5B3A8A",
}
SEQ_COLORS = ["#2E6B4F","#1B4F8A","#C45C1A","#B8860B","#6B7A8D","#5B3A8A","#B03A2E"]

# ── DATA LOADER ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados…")
def load_data(path: str):
    xls = pd.ExcelFile(path)
    dfs = {}
    for sheet in xls.sheet_names:
        if sheet == "LEIA-ME":
            continue
        dfs[sheet] = pd.read_excel(xls, sheet_name=sheet)

    # Fix insumos duplicate col
    ins = dfs.get("INSUMOS", pd.DataFrame())
    if not ins.empty:
        cols = list(ins.columns)
        if cols.count("Safra") == 2:
            idx = [i for i,c in enumerate(cols) if c == "Safra"]
            cols[idx[1]] = "Safra_ERP"
            ins.columns = cols
        dfs["INSUMOS"] = ins

    # Dívida — keep only real records
    div = dfs.get("DIVIDA", pd.DataFrame())
    if not div.empty:
        div = div[div["Moeda"].isin(["R$","US$"]) & div["Vencimento"].notna()].copy()
        dfs["DIVIDA"] = div

    # Exposição — drop dupes
    exp = dfs.get("EXPOSICAO_MENSAL", pd.DataFrame())
    if not exp.empty:
        exp = exp.drop_duplicates(subset=["Safra_Base","Safra","Data","Metrica"])
        dfs["EXPOSICAO_MENSAL"] = exp

    return dfs

def fmt_usd(v, decimals=2):
    if pd.isna(v): return "—"
    neg = v < 0
    s = f"U$ {abs(v/1e6):,.{decimals}f}M"
    return f"-{s}" if neg else s

def fmt_k(v):
    if pd.isna(v): return "—"
    return f"{abs(v/1e3):,.0f}k sc"

def kpi(label, value, sub="", cls=""):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {cls}">{value}</div>
        {"<div class='kpi-sub'>"+sub+"</div>" if sub else ""}
    </div>"""

def badge(text, tipo="long"):
    return f'<span class="badge badge-{tipo}">{text}</span>'

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="sidebar-logo">🌱 GFRT</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">FX Exposure Dashboard</div>', unsafe_allow_html=True)
    st.markdown("---")

    excel_path = st.text_input(
        "Caminho do arquivo Excel",
        value="GFRT_FX_Exposure_PowerBI.xlsx",
        help="Informe o caminho completo ou relativo do Excel gerado."
    )

    try:
        dfs = load_data(excel_path)
        st.success("✓ Dados carregados", icon=None)
    except Exception as e:
        st.error(f"Arquivo não encontrado.\n\n`{e}`")
        st.stop()

    st.markdown("---")
    page = st.radio(
        "Página",
        ["Balanço FX", "Exposição Mensal", "Soja / Comercialização",
         "Insumos USD", "Dívida e Terras"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.markdown('<div style="font-size:10px;color:#3C4D5E">Grupo Fernando Ribas Taques<br>Planejamento Financeiro</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — BALANÇO FX
# ══════════════════════════════════════════════════════════════════════════════
if page == "Balanço FX":
    df_bal = dfs.get("BALANCO_FX", pd.DataFrame())

    st.markdown("""
    <div class="page-header">
        <h2>Balanço FX</h2>
        <p>Posição consolidada ativo × passivo em dólar por safra</p>
    </div>""", unsafe_allow_html=True)

    safras = sorted(df_bal["Safra"].dropna().unique()) if not df_bal.empty else ["25/26"]
    col_saf, _ = st.columns([2, 8])
    with col_saf:
        safra_sel = st.selectbox("Safra", safras, label_visibility="collapsed")

    d = df_bal[df_bal["Safra"] == safra_sel]
    ativo   = d[d["Categoria"] == "Ativo"]["Valor_USD"].sum()
    passivo = d[d["Categoria"] == "Passivo"]["Valor_USD"].sum()
    liquido = ativo + passivo
    cobert  = (ativo / abs(passivo) * 100) if passivo != 0 else 0
    status  = "LONG" if liquido >= 0 else "SHORT"
    cls_liq = "kpi-pos" if liquido >= 0 else "kpi-neg"

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi("Total ativo", fmt_usd(ativo), "Soja vendida + a vender", "kpi-pos"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("Total passivo", fmt_usd(passivo), "Insumos + terras + dívida", "kpi-neg"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("Saldo líquido", fmt_usd(liquido), badge(status, status.lower()), cls_liq), unsafe_allow_html=True)
    with c4: st.markdown(kpi("Cobertura ativo/passivo", f"{cobert:.0f}%", "Ativo ÷ passivo"), unsafe_allow_html=True)

    st.markdown('<div class="sect">Composição</div>', unsafe_allow_html=True)

    col_l, col_r = st.columns([1, 1])

    # Donut passivo
    with col_l:
        pass_d = d[d["Categoria"] == "Passivo"].copy()
        pass_d["Valor_Abs"] = pass_d["Valor_USD"].abs()
        fig = go.Figure(go.Pie(
            labels=pass_d["Subcategoria"],
            values=pass_d["Valor_Abs"],
            hole=.55,
            marker_colors=SEQ_COLORS,
            textinfo="label+percent",
            textfont_size=12,
        ))
        fig.update_layout(
            title=dict(text="Composição do passivo", font_size=13, x=0),
            showlegend=False,
            margin=dict(t=40, b=10, l=10, r=10),
            height=300,
            paper_bgcolor="white",
            plot_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)

    # Waterfall
    with col_r:
        cats = list(d["Subcategoria"])
        vals = list(d["Valor_USD"])
        labels = cats + ["Líquido"]
        measures = ["relative"] * len(cats) + ["total"]
        values = vals + [liquido]
        colors = ["#2E6B4F" if v >= 0 else "#B03A2E" for v in values]
        colors[-1] = "#1B4F8A"

        fig2 = go.Figure(go.Waterfall(
            name="",
            orientation="v",
            measure=measures,
            x=labels,
            y=values,
            connector={"line": {"color": "#ECEAE5"}},
            increasing={"marker": {"color": "#2E6B4F"}},
            decreasing={"marker": {"color": "#B03A2E"}},
            totals={"marker": {"color": "#1B4F8A"}},
            texttemplate="%{y:,.0f}",
            textfont_size=10,
        ))
        fig2.update_layout(
            title=dict(text="Waterfall ativo × passivo", font_size=13, x=0),
            showlegend=False,
            margin=dict(t=40, b=10, l=10, r=10),
            height=300,
            paper_bgcolor="white",
            plot_bgcolor="white",
            yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
            xaxis=dict(tickfont_size=11),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="sect">Detalhamento por subcategoria</div>', unsafe_allow_html=True)
    col_a, col_p = st.columns(2)

    with col_a:
        ativo_d = d[d["Categoria"] == "Ativo"][["Subcategoria","Valor_USD"]].copy()
        ativo_d.columns = ["Subcategoria","Valor USD"]
        ativo_d["Valor USD"] = ativo_d["Valor USD"].apply(fmt_usd)
        st.markdown("**Ativo**")
        st.dataframe(ativo_d, hide_index=True, use_container_width=True)

    with col_p:
        pass_d2 = d[d["Categoria"] == "Passivo"][["Subcategoria","Valor_USD"]].copy()
        pass_d2.columns = ["Subcategoria","Valor USD"]
        pass_d2["Valor USD"] = pass_d2["Valor USD"].apply(fmt_usd)
        st.markdown("**Passivo**")
        st.dataframe(pass_d2, hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — EXPOSIÇÃO MENSAL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Exposição Mensal":
    df_exp = dfs.get("EXPOSICAO_MENSAL", pd.DataFrame())

    st.markdown("""
    <div class="page-header">
        <h2>Exposição Mensal</h2>
        <p>Fluxo mensal de entradas e saídas em dólar com posição acumulada</p>
    </div>""", unsafe_allow_html=True)

    safras_exp = sorted(df_exp["Safra_Base"].dropna().unique()) if not df_exp.empty else []
    col_sf, _ = st.columns([2, 8])
    with col_sf:
        safra_e = st.selectbox("Safra", safras_exp, label_visibility="collapsed")

    de = df_exp[df_exp["Safra_Base"] == safra_e].copy()
    de["Data"] = pd.to_datetime(de["Data"])

    entradas = de[de["Metrica"] == "Venda Soja"]["Valor_USD"].sum()
    saidas   = de[de["Metrica"] == "SAÍDAS"]["Valor_USD"].sum()
    fx_total = de[de["Metrica"] == "FX EXPOUSURE"]["Valor_USD"].sum()
    ins_t    = de[de["Metrica"] == "Insumos"]["Valor_USD"].sum()
    div_t    = de[de["Metrica"] == "Dívida"]["Valor_USD"].sum()
    status_e = "LONG" if fx_total >= 0 else "SHORT"

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi("Entradas soja", fmt_usd(entradas), "Recebimentos no período", "kpi-pos"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("FX exposure", fmt_usd(fx_total), badge(status_e, status_e.lower()), "kpi-pos" if fx_total>=0 else "kpi-neg"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("Saídas insumos", fmt_usd(ins_t), "Vencimentos período", "kpi-neg"), unsafe_allow_html=True)
    with c4: st.markdown(kpi("Saídas dívida", fmt_usd(div_t), "Amortizações período", "kpi-neg"), unsafe_allow_html=True)

    st.markdown('<div class="sect">Entradas × saídas por mês</div>', unsafe_allow_html=True)

    # Pivot por mês e métrica
    pivot_metr = ["Venda Soja","Insumos","Terras","Dívida","Outros"]
    cores_metr = {"Venda Soja": "#2E6B4F","Insumos": "#C45C1A","Terras": "#1B4F8A","Dívida": "#B8860B","Outros": "#6B7A8D"}

    df_pivot = de[de["Metrica"].isin(pivot_metr)].copy()
    df_pivot["Data"] = pd.to_datetime(df_pivot["Data"]).dt.strftime("%b/%y")
    df_pivot["Valor_USD"] = df_pivot.apply(
        lambda r: r["Valor_USD"] if r["Metrica"] == "Venda Soja" else -abs(r["Valor_USD"]), axis=1
    )

    fig3 = go.Figure()
    for met in pivot_metr:
        sub = df_pivot[df_pivot["Metrica"] == met]
        if sub.empty: continue
        fig3.add_trace(go.Bar(
            x=sub["Data"], y=sub["Valor_USD"],
            name=met,
            marker_color=cores_metr.get(met, "#888"),
        ))
    fig3.update_layout(
        barmode="relative",
        height=300,
        paper_bgcolor="white", plot_bgcolor="white",
        margin=dict(t=10, b=30, l=10, r=10),
        legend=dict(orientation="h", y=1.08, font_size=11),
        yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
        xaxis=dict(tickfont_size=11),
    )
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown('<div class="sect">FX exposure acumulado</div>', unsafe_allow_html=True)

    df_fx = de[de["Metrica"] == "FX EXPOUSURE"].sort_values("Data").copy()
    df_fx["Acumulado"] = df_fx["Valor_USD"].cumsum()
    df_fx["Mes"] = pd.to_datetime(df_fx["Data"]).dt.strftime("%b/%y")

    fig4 = go.Figure()
    fig4.add_trace(go.Scatter(
        x=df_fx["Mes"], y=df_fx["Acumulado"],
        mode="lines+markers",
        line=dict(color="#B03A2E", width=2.5),
        marker=dict(size=7, color="#B03A2E"),
        fill="tozeroy",
        fillcolor="rgba(176,58,46,0.07)",
        name="FX acumulado",
    ))
    fig4.add_hline(y=0, line_dash="dot", line_color="#AAAAAA", line_width=1)
    fig4.update_layout(
        height=240,
        paper_bgcolor="white", plot_bgcolor="white",
        margin=dict(t=10, b=30, l=10, r=10),
        showlegend=False,
        yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
        xaxis=dict(tickfont_size=11),
    )
    st.plotly_chart(fig4, use_container_width=True)

    st.markdown('<div class="sect">Detalhe mensal</div>', unsafe_allow_html=True)
    df_tab = de[de["Metrica"] == "FX EXPOUSURE"].sort_values("Data").copy()
    df_tab["Mês"] = pd.to_datetime(df_tab["Data"]).dt.strftime("%b/%y")
    df_tab["FX Exposure"] = df_tab["Valor_USD"].apply(fmt_usd)
    df_tab["Status"] = df_tab["Valor_USD"].apply(lambda v: "LONG" if v >= 0 else "SHORT")
    st.dataframe(df_tab[["Mês","FX Exposure","Status"]].reset_index(drop=True),
                 hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — SOJA / COMERCIALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Soja / Comercialização":
    df_prod = dfs.get("PRODUCAO", pd.DataFrame())
    df_ctt  = dfs.get("CONTRATOS_SOJA", pd.DataFrame())

    st.markdown("""
    <div class="page-header">
        <h2>Soja / Comercialização</h2>
        <p>Produção projetada, contratos fechados e posição de comercialização</p>
    </div>""", unsafe_allow_html=True)

    safras_s = sorted(df_prod["Safra"].dropna().unique()) if not df_prod.empty else []
    col_sf, _ = st.columns([2,8])
    with col_sf:
        safra_s = st.selectbox("Safra", safras_s, label_visibility="collapsed")

    dp = df_prod[df_prod["Safra"] == safra_s]
    dc = df_ctt[df_ctt["Safra"] == safra_s] if not df_ctt.empty else pd.DataFrame()

    total_prod  = dp["Total_Sacas"].sum()
    vend_usd    = dc[dc["Moeda"] == "Dólar"]["Qtd_Sacas"].sum() if not dc.empty else 0
    vend_brl    = dc[dc["Moeda"] == "Reais"]["Qtd_Sacas"].sum() if not dc.empty else 0
    total_vend  = vend_usd + vend_brl
    a_vender    = total_prod - total_vend
    pct_vend    = total_vend / total_prod * 100 if total_prod else 0
    val_usd     = dc[dc["Moeda"] == "Dólar"]["Valor_Total"].sum() if not dc.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi("Total produzido", fmt_k(total_prod), f"{dp['Area_Ha'].sum():,.0f} ha"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("Sacas vendidas", fmt_k(total_vend), f"{pct_vend:.1f}% da produção", "kpi-pos"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("A vender", fmt_k(a_vender), f"{100-pct_vend:.1f}% descoberto", "kpi-neg"), unsafe_allow_html=True)
    with c4: st.markdown(kpi("Receita USD", fmt_usd(val_usd), f"{vend_usd:,.0f} sc em USD"), unsafe_allow_html=True)

    st.markdown('<div class="sect">Produção e comercialização por fazenda</div>', unsafe_allow_html=True)
    col_l, col_r = st.columns([1.2, 1])

    with col_l:
        # Grouped bar por fazenda
        fazendas = dp["Fazenda"].tolist()
        prod_vals = dp.set_index("Fazenda")["Total_Sacas"]

        vend_por_faz = dc.groupby("Fazenda")["Qtd_Sacas"].sum() if not dc.empty else pd.Series(dtype=float)

        fig5 = go.Figure()
        fig5.add_trace(go.Bar(
            x=fazendas,
            y=[prod_vals.get(f, 0) for f in fazendas],
            name="Produzido",
            marker_color="#ECEAE5",
            text=[f"{v/1e3:.0f}k" for v in [prod_vals.get(f,0) for f in fazendas]],
            textposition="outside",
        ))
        fig5.add_trace(go.Bar(
            x=fazendas,
            y=[vend_por_faz.get(f, 0) for f in fazendas],
            name="Vendido",
            marker_color="#2E6B4F",
            text=[f"{v/1e3:.0f}k" for v in [vend_por_faz.get(f,0) for f in fazendas]],
            textposition="outside",
        ))
        fig5.update_layout(
            barmode="overlay",
            height=280,
            paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(t=10, b=10, l=10, r=10),
            legend=dict(orientation="h", y=1.1, font_size=11),
            yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
            xaxis=dict(tickfont_size=11),
        )
        st.plotly_chart(fig5, use_container_width=True)

    with col_r:
        # Donut USD x BRL x a vender
        fig6 = go.Figure(go.Pie(
            labels=["Vendido USD","Vendido R$","A vender"],
            values=[max(vend_usd,0), max(vend_brl,0), max(a_vender,0)],
            hole=.55,
            marker_colors=["#1B4F8A","#2E6B4F","#ECEAE5"],
            textinfo="label+percent",
            textfont_size=11,
        ))
        fig6.update_layout(
            showlegend=False,
            height=280,
            paper_bgcolor="white",
            margin=dict(t=10, b=10, l=10, r=10),
        )
        st.plotly_chart(fig6, use_container_width=True)

    st.markdown('<div class="sect">Contratos fechados</div>', unsafe_allow_html=True)
    if not dc.empty:
        show_cols = ["Fazenda","Trader","Mediador","Qtd_Sacas","Moeda","Valor_Saca","Valor_Total","Data_Recebimento","Saldo_Receber"]
        show_cols = [c for c in show_cols if c in dc.columns]
        dc_disp = dc[show_cols].copy()
        if "Data_Recebimento" in dc_disp.columns:
            dc_disp["Data_Recebimento"] = pd.to_datetime(dc_disp["Data_Recebimento"]).dt.strftime("%d/%m/%Y")
        for col in ["Qtd_Sacas","Valor_Total","Valor_Saca","Saldo_Receber"]:
            if col in dc_disp.columns:
                dc_disp[col] = dc_disp[col].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "—")
        st.dataframe(dc_disp.rename(columns={
            "Qtd_Sacas":"Sacas","Valor_Saca":"Preço/sc","Valor_Total":"Total",
            "Data_Recebimento":"Recebimento","Saldo_Receber":"Saldo a Rec."
        }), hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 4 — INSUMOS USD
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Insumos USD":
    df_ins = dfs.get("INSUMOS", pd.DataFrame())
    df_prod = dfs.get("PRODUCAO", pd.DataFrame())

    st.markdown("""
    <div class="page-header">
        <h2>Insumos USD</h2>
        <p>Compras de insumos em dólar — vencimentos, fornecedores e fazendas</p>
    </div>""", unsafe_allow_html=True)

    if not df_ins.empty:
        cols = list(df_ins.columns)
        if cols.count("Safra") == 2:
            idx = [i for i,c in enumerate(cols) if c=="Safra"]
            cols[idx[1]] = "Safra_ERP"
            df_ins.columns = cols

    safras_i = sorted(df_ins["Safra"].dropna().unique()) if not df_ins.empty else []
    col_sf, col_tipo, _ = st.columns([1.5, 2, 6])
    with col_sf:
        safra_i = st.selectbox("Safra", safras_i, label_visibility="collapsed")
    di = df_ins[df_ins["Safra"] == safra_i].copy()

    tipos = ["Todos"] + sorted(di["Tipo"].dropna().unique().tolist())
    with col_tipo:
        tipo_sel = st.selectbox("Tipo", tipos, label_visibility="collapsed")
    if tipo_sel != "Todos":
        di = di[di["Tipo"] == tipo_sel]

    total_usd = di["A_pagar_Dolar"].sum()
    total_sc  = di["Sacas"].sum()
    prod_safra = df_prod[df_prod["Safra"]==safra_i]["Total_Sacas"].sum()
    pct_prod  = total_sc / prod_safra * 100 if prod_safra else 0
    n_forn    = di["Fornecedor"].nunique()

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi("Total insumos USD", fmt_usd(total_usd), "Dólar + cessões"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("Em sacas equiv.", fmt_k(total_sc), "Custo convertido em soja"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("% da produção", f"{pct_prod:.1f}%", "Insumos ÷ sacas produzidas"), unsafe_allow_html=True)
    with c4: st.markdown(kpi("Fornecedores", str(n_forn), "Ativos no período"), unsafe_allow_html=True)

    st.markdown('<div class="sect">Análise</div>', unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns(3)

    # Donut por tipo
    with col_l:
        tipo_g = di.groupby("Tipo")["A_pagar_Dolar"].sum().sort_values(ascending=False)
        fig7 = go.Figure(go.Pie(
            labels=tipo_g.index, values=tipo_g.values,
            hole=.55, marker_colors=SEQ_COLORS,
            textinfo="label+percent", textfont_size=11,
        ))
        fig7.update_layout(title=dict(text="Por tipo", font_size=13, x=0),
            showlegend=False, height=280, paper_bgcolor="white",
            margin=dict(t=40,b=5,l=5,r=5))
        st.plotly_chart(fig7, use_container_width=True)

    # Bar fornecedor
    with col_m:
        forn_g = di.groupby("Fornecedor")["A_pagar_Dolar"].sum().sort_values(ascending=False).head(8)
        fig8 = go.Figure(go.Bar(
            x=forn_g.values, y=forn_g.index,
            orientation="h",
            marker_color=COLORS["azul"],
            text=[fmt_usd(v,1) for v in forn_g.values],
            textposition="auto",
        ))
        fig8.update_layout(title=dict(text="Top fornecedores", font_size=13, x=0),
            height=280, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(t=40,b=5,l=5,r=5),
            xaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
            yaxis=dict(tickfont_size=10),
            showlegend=False)
        st.plotly_chart(fig8, use_container_width=True)

    # Bar fazenda
    with col_r:
        faz_g = di.groupby("Fazenda")["A_pagar_Dolar"].sum().sort_values(ascending=False)
        fig9 = go.Figure(go.Bar(
            x=faz_g.index, y=faz_g.values,
            marker_color=COLORS["verde"],
            text=[fmt_usd(v,1) for v in faz_g.values],
            textposition="outside",
        ))
        fig9.update_layout(title=dict(text="Por fazenda", font_size=13, x=0),
            height=280, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(t=40,b=5,l=5,r=5),
            yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
            xaxis=dict(tickfont_size=10),
            showlegend=False)
        st.plotly_chart(fig9, use_container_width=True)

    st.markdown('<div class="sect">Cronograma de vencimentos (USD)</div>', unsafe_allow_html=True)
    di_v = di[di["Vencimento"].notna() & (di["A_pagar_Dolar"] > 0)].copy()
    di_v["Vencimento"] = pd.to_datetime(di_v["Vencimento"])
    di_v["Mês"] = di_v["Vencimento"].dt.to_period("M").astype(str)
    vcto_g = di_v.groupby(["Mês","Tipo"])["A_pagar_Dolar"].sum().reset_index()

    if not vcto_g.empty:
        fig10 = px.bar(vcto_g, x="Mês", y="A_pagar_Dolar", color="Tipo",
                       color_discrete_sequence=SEQ_COLORS,
                       labels={"A_pagar_Dolar":"USD","Mês":"Vencimento"})
        fig10.update_layout(height=260, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(t=10,b=10,l=10,r=10),
            legend=dict(orientation="h", y=1.08, font_size=11),
            yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"))
        st.plotly_chart(fig10, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 5 — DÍVIDA E TERRAS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Dívida e Terras":
    df_div = dfs.get("DIVIDA", pd.DataFrame())
    df_ter = dfs.get("TERRAS", pd.DataFrame())

    st.markdown("""
    <div class="page-header">
        <h2>Dívida e Terras</h2>
        <p>Passivos financeiros bancários e arrendamentos indexados à soja</p>
    </div>""", unsafe_allow_html=True)

    div_usd = df_div[df_div["Moeda"] == "US$"].copy() if not df_div.empty else pd.DataFrame()
    div_brl = df_div[df_div["Moeda"] == "R$"].copy()  if not df_div.empty else pd.DataFrame()

    total_div_usd = div_usd["Valor"].sum() if not div_usd.empty else 0
    total_div_brl = div_brl["Valor"].sum() if not div_brl.empty else 0
    total_ter     = df_ter["Valor_Total_USD"].sum() if not df_ter.empty else 0

    # Vence em 12 meses
    div_usd_cp = div_usd.copy()
    if not div_usd_cp.empty:
        div_usd_cp["Vencimento"] = pd.to_datetime(div_usd_cp["Vencimento"])
        hoje = pd.Timestamp.today()
        vence_12m = div_usd_cp[
            (div_usd_cp["Vencimento"] >= hoje) &
            (div_usd_cp["Vencimento"] <= hoje + pd.DateOffset(months=12))
        ]["Valor"].sum()
    else:
        vence_12m = 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi("Dívida bancária USD", fmt_usd(total_div_usd), "5 bancos", "kpi-neg"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("Vence em 12 meses", fmt_usd(vence_12m), "Curto prazo USD", "kpi-neg"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("Terras USD total", fmt_usd(total_ter), "Renascer — indexado soja", "kpi-neg"), unsafe_allow_html=True)
    with c4: st.markdown(kpi("Dívida BRL total", f"R$ {total_div_brl/1e6:.1f}M", "Incluindo juros"), unsafe_allow_html=True)

    st.markdown('<div class="sect">Dívida bancária USD</div>', unsafe_allow_html=True)
    col_l, col_r = st.columns([1, 1])

    with col_l:
        if not div_usd.empty:
            banco_g = div_usd.groupby("Cliente_Fornecedor")["Valor"].sum().sort_values(ascending=False)
            # Shorten names
            banco_labels = [n[:28] + "…" if len(n) > 28 else n for n in banco_g.index]
            fig11 = go.Figure(go.Bar(
                x=banco_g.values, y=banco_labels,
                orientation="h",
                marker_color=COLORS["azul"],
                text=[fmt_usd(v,1) for v in banco_g.values],
                textposition="auto",
            ))
            fig11.update_layout(
                title=dict(text="Concentração por banco", font_size=13, x=0),
                height=280, paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(t=40,b=5,l=5,r=5),
                xaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
                yaxis=dict(tickfont_size=10),
                showlegend=False,
            )
            st.plotly_chart(fig11, use_container_width=True)

    with col_r:
        if not div_usd.empty:
            tipo_g = div_usd.groupby("Tipo")["Valor"].sum()
            fig12 = go.Figure(go.Pie(
                labels=tipo_g.index, values=tipo_g.values,
                hole=.55,
                marker_colors=["#1B4F8A","#B5D4F4"],
                textinfo="label+percent", textfont_size=11,
            ))
            fig12.update_layout(
                title=dict(text="Principal × juros", font_size=13, x=0),
                showlegend=False, height=280, paper_bgcolor="white",
                margin=dict(t=40,b=5,l=5,r=5),
            )
            st.plotly_chart(fig12, use_container_width=True)

    st.markdown('<div class="sect">Cronograma de amortização por ano</div>', unsafe_allow_html=True)

    if not div_usd.empty:
        div_usd["Vencimento"] = pd.to_datetime(div_usd["Vencimento"])
        div_usd["Ano"] = div_usd["Vencimento"].dt.year
        ano_g = div_usd.groupby("Ano")["Valor"].sum().reset_index()

        ter_ano = pd.DataFrame()
        if not df_ter.empty:
            df_ter2 = df_ter.copy()
            df_ter2["Vencimento"] = pd.to_datetime(df_ter2["Vencimento"])
            df_ter2["Ano"] = df_ter2["Vencimento"].dt.year
            ter_ano = df_ter2.groupby("Ano")["Valor_Total_USD"].sum().reset_index()

        anos = sorted(set(ano_g["Ano"].tolist() + (ter_ano["Ano"].tolist() if not ter_ano.empty else [])))

        fig13 = go.Figure()
        fig13.add_trace(go.Bar(
            x=anos,
            y=[ano_g[ano_g["Ano"]==a]["Valor"].sum() for a in anos],
            name="Dívida bancária",
            marker_color=COLORS["laranja"],
        ))
        if not ter_ano.empty:
            fig13.add_trace(go.Bar(
                x=anos,
                y=[ter_ano[ter_ano["Ano"]==a]["Valor_Total_USD"].sum() for a in anos],
                name="Terras",
                marker_color=COLORS["azul"],
            ))
        fig13.update_layout(
            barmode="stack",
            height=260, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(t=10,b=10,l=10,r=10),
            legend=dict(orientation="h", y=1.1, font_size=11),
            yaxis=dict(tickformat=",.0f", gridcolor="#F0EFEB"),
        )
        st.plotly_chart(fig13, use_container_width=True)

    st.markdown('<div class="sect">Parcelas de arrendamento — Renascer</div>', unsafe_allow_html=True)
    if not df_ter.empty:
        ter_disp = df_ter[["Propriedade","Vencimento","Nr_Sacas","Valor_Saca_USD","Valor_Total_USD","Safra"]].copy()
        ter_disp["Vencimento"] = pd.to_datetime(ter_disp["Vencimento"]).dt.strftime("%d/%m/%Y")
        ter_disp["Nr_Sacas"]   = ter_disp["Nr_Sacas"].apply(lambda v: f"{v:,.0f}")
        ter_disp["Valor_Saca_USD"]  = ter_disp["Valor_Saca_USD"].apply(lambda v: f"U$ {v:.2f}")
        ter_disp["Valor_Total_USD"] = ter_disp["Valor_Total_USD"].apply(fmt_usd)
        ter_disp.columns = ["Propriedade","Vencimento","Sacas","USD/saca","Total USD","Safra"]
        st.dataframe(ter_disp.reset_index(drop=True), hide_index=True, use_container_width=True)
