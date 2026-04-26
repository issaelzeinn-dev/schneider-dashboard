import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd

st.set_page_config(page_title="Schneider Electric - Financial Dashboard", layout="wide", page_icon="⚡")

# ── THEME ──────────────────────────────────────────────────────────
NAVY   = "#1F497D"
TEAL   = "#0A6E6E"
GREEN  = "#1A7A3C"
AMBER  = "#C47800"
RED    = "#C0392B"
LGRAY  = "#F5F7FA"
MGRAY  = "#E8ECF0"
DGRAY  = "#5A6473"
WHITE  = "#FFFFFF"
YEARS  = [2021, 2022, 2023, 2024, 2025]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main { background: #F8FAFB; }
.block-container { padding: 1.2rem 2rem 2rem; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #E8ECF0; border-radius: 10px; padding: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 18px; font-weight: 500; font-size: 13px; }
.stTabs [aria-selected="true"] { background: #1F497D !important; color: white !important; }
div[data-testid="metric-container"] { background: white; border-radius: 10px; padding: 14px; border: 1px solid #E2E8F0; }
</style>
""", unsafe_allow_html=True)

# ── DATA ───────────────────────────────────────────────────────────
@st.cache_data
def load():
    path = "Schneider_Dashboard.xlsx"

    r  = pd.read_excel(path, sheet_name="Ratios",              header=None, engine="openpyxl")
    i  = pd.read_excel(path, sheet_name="Income Statement",    header=None, engine="openpyxl")
    cf = pd.read_excel(path, sheet_name="Cash Flow Statement", header=None, engine="openpyxl")
    bs = pd.read_excel(path, sheet_name="Balance Sheet",       header=None, engine="openpyxl")

    def to_num(v):
        try: return float(v)
        except: return None

    # NEW column positions (FY2021 -> FY2025)
    IC = [28, 25, 22, 19, 16]   # Income Statement
    RC = [9,  8,  7,  6,  5]    # Ratios
    CC = [26, 24, 22, 20, 18]   # Cash Flow Statement

    def irow(df, c): return [to_num(df.iloc[c, x]) for x in IC]
    def rrow(df, c): return [to_num(df.iloc[c, x]) for x in RC]
    def crow(df, c): return [to_num(df.iloc[c, x]) for x in CC]

    d = {}

    # Income Statement - new row positions confirmed
    d["revenue"]      = irow(i, 3)
    d["cogs"]         = irow(i, 4)
    d["gross_profit"] = irow(i, 5)
    d["ebita"]        = irow(i, 11)
    d["ebitda"]       = irow(i, 12)
    d["ebit"]         = irow(i, 14)
    d["net_income"]   = irow(i, 24)
    d["int_expense"]  = irow(i, 16)

    # Ratios - new row positions confirmed
    d["roe"]          = rrow(r, 3)
    d["ros"]          = rrow(r, 4)
    d["asset_to"]     = rrow(r, 5)
    d["eq_mult"]      = rrow(r, 6)
    d["gross_m"]      = rrow(r, 10)
    d["ebita_m"]      = rrow(r, 13)
    d["ebitda_m"]     = rrow(r, 14)
    d["ebit_m"]       = rrow(r, 15)
    d["dso"]          = rrow(r, 17)
    d["dio"]          = rrow(r, 18)
    d["dpo"]          = rrow(r, 19)
    d["ccc"]          = rrow(r, 20)
    d["current_r"]    = rrow(r, 24)
    d["quick_r"]      = rrow(r, 25)
    d["cash_r"]       = rrow(r, 26)
    d["de_ratio"]     = rrow(r, 28)
    d["debt_ratio"]   = rrow(r, 29)
    d["int_cov"]      = rrow(r, 30)
    d["nd_ebita"]     = rrow(r, 31)
    d["nd_ebitda"]    = rrow(r, 32)
    d["payout"]       = rrow(r, 34)
    d["div_yield"]    = rrow(r, 35)
    d["eps"]          = rrow(r, 36)
    d["pe"]           = rrow(r, 37)
    d["pb"]           = rrow(r, 38)
    d["net_debt"]     = rrow(r, 39)
    d["ev"]           = rrow(r, 40)
    d["ev_ebita"]     = rrow(r, 41)
    d["roce"]         = rrow(r, 43)

    # Derived ratios - with safe division
    d["int_cov_ebd"]  = [to_num(eb)/abs(to_num(ie)) if to_num(ie) else None
                          for eb,ie in zip(d["ebitda"], d["int_expense"])]

    # Cash Flow Statement - same rows, new columns
    d["cfo"]     = crow(cf, 19)
    d["capex"]   = crow(cf, 23)
    d["cfi"]     = crow(cf, 28)
    d["cff"]     = crow(cf, 38)
    d["divs"]    = crow(cf, 36)
    d["buyback"] = crow(cf, 31)
    d["ma"]      = crow(cf, 24)

    def sdiv(a, b):
        try:
            a, b = float(a), float(b)
            return a / b if b else None
        except: return None

    d["fcf"]     = [c + cap if c is not None and cap is not None else None
                    for c,cap in zip(d["cfo"], d["capex"])]
    d["fcf_m"]   = [sdiv(f, rev) for f,rev in zip(d["fcf"],     d["revenue"])]
    d["fcf_ni"]  = [sdiv(f, ni)  for f,ni  in zip(d["fcf"],     d["net_income"])]
    d["cfo_ebd"] = [sdiv(c, e)   for c,e   in zip(d["cfo"],     d["ebitda"])]
    d["capex_r"] = [sdiv(abs(cap) if cap else None, rev)
                    for cap,rev in zip(d["capex"], d["revenue"])]
    dda = [sdiv(e - eb, 1) if e is not None and eb is not None else None
           for e,eb in zip(d["ebitda"], d["ebita"])]
    d["capex_da"]= [sdiv(abs(cap) if cap else None, da)
                    for cap,da in zip(d["capex"], dda)]

    # Balance Sheet - new col positions: FY2025=15, FY2024=18
    def bval(row, col): return to_num(bs.iloc[row, col])
    d["intangibles"]  = [bval(2,  15), bval(2,  18)]
    d["ppe"]          = [bval(3,  15), bval(3,  18)]
    d["oth_nca"]      = [bval(4,  15), bval(4,  18)]
    d["inventories"]  = [bval(6,  15), bval(6,  18)]
    d["receivables"]  = [bval(7,  15), bval(7,  18)]
    d["cash"]         = [bval(9,  15), bval(9,  18)]
    d["total_assets"] = [bval(11, 15), bval(11, 18)]
    d["equity"]       = [bval(16, 15), bval(16, 18)]
    d["lt_debt"]      = [bval(17, 15), bval(17, 18)]
    d["st_debt"]      = [bval(22, 15), bval(22, 18)]

    return d

d = load()

def safe(val, fmt="pct", dec=1):
    try:
        v = float(val)
        if fmt == "pct":   return f"{v*100:.{dec}f}%"
        if fmt == "x":     return f"{v:.{dec}f}x"
        if fmt == "eur":   return f"{v:,.0f}"
        if fmt == "num":   return f"{v:.{dec}f}"
        if fmt == "days":  return f"{v:.1f}d"
        return str(v)
    except: return "--"

def delta_str(curr, prev, inverse=False):
    try:
        d = float(curr) - float(prev)
        symbol = "▲" if d > 0 else "▼"
        color = GREEN if (d > 0) != inverse else RED
        return f'<span style="color:{color};font-size:12px">{symbol} {abs(d)*100:.1f}pp</span>'
    except: return ""

def delta_abs(curr, prev, inverse=False):
    try:
        d = float(curr) - float(prev)
        symbol = "▲" if d > 0 else "▼"
        color = GREEN if (d > 0) != inverse else RED
        pct = (float(curr)-float(prev))/abs(float(prev))*100
        return f'<span style="color:{color};font-size:12px">{symbol} {pct:.1f}%</span>'
    except: return ""

def kpi_card(label, value, delta_html="", badge=None):
    badge_html = f'<span style="background:#E8F4FD;color:#0A6E6E;font-size:9px;font-weight:600;padding:2px 6px;border-radius:4px;margin-left:6px">{badge}</span>' if badge else ""
    st.markdown(f"""
    <div style="background:white;border-radius:10px;padding:16px;border:1px solid #E2E8F0;height:100%">
      <div style="font-size:10px;font-weight:600;color:{DGRAY};text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px">
        {label}{badge_html}
      </div>
      <div style="font-size:26px;font-weight:700;color:{NAVY};line-height:1.1">{value}</div>
      <div style="margin-top:4px">{delta_html}</div>
    </div>""", unsafe_allow_html=True)

def section(title, color=NAVY):
    st.markdown(f"""<div style="background:{color};color:white;padding:7px 14px;border-radius:6px;
    font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin:18px 0 10px">
    {title}</div>""", unsafe_allow_html=True)

def line_chart(title, series_dict, y_fmt="pct", height=280):
    fig = go.Figure()
    colors = [TEAL, NAVY, GREEN, AMBER, RED]
    for i, (name, vals) in enumerate(series_dict.items()):
        yv = [v*100 if y_fmt=="pct" and v is not None else v for v in vals]
        fig.add_trace(go.Scatter(x=YEARS, y=yv, name=name,
            line=dict(color=colors[i % len(colors)], width=2.5),
            mode="lines+markers", marker=dict(size=6)))
    ytitle = "%" if y_fmt=="pct" else "EURm"
    fig.update_layout(height=height, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.2, font=dict(size=10)),
        yaxis=dict(ticksuffix="%" if y_fmt=="pct" else "", gridcolor="#F0F0F0", tickfont=dict(size=10)),
        xaxis=dict(tickfont=dict(size=10), dtick=1))
    return fig

def bar_chart(title, series_dict, height=280, stack=False):
    fig = go.Figure()
    colors = [TEAL, NAVY, GREEN, AMBER, RED, "#8E44AD"]
    bt = "stack" if stack else "group"
    for i, (name, vals) in enumerate(series_dict.items()):
        c = colors[i % len(colors)]
        fig.add_trace(go.Bar(x=YEARS, y=vals, name=name,
            marker_color=c, opacity=0.88))
    fig.update_layout(barmode=bt, height=height, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.2, font=dict(size=10)),
        yaxis=dict(gridcolor="#F0F0F0", tickfont=dict(size=10)),
        xaxis=dict(tickfont=dict(size=10), dtick=1))
    return fig

def donut(labels, values, title, colors_list):
    fig = go.Figure(go.Pie(labels=labels, values=[abs(v) for v in values],
        hole=0.62, marker_colors=colors_list,
        textinfo="percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:,.0f} EURm<extra></extra>"))
    fig.update_layout(height=220, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=True, legend=dict(font=dict(size=9), orientation="h", y=-0.1))
    return fig

def gauge(value, title, min_v, max_v, thresholds, fmt="pct"):
    display = value*100 if fmt=="pct" else value
    vmin = min_v*100 if fmt=="pct" else min_v
    vmax = max_v*100 if fmt=="pct" else max_v
    steps = []
    prev = vmin
    cols = [RED, AMBER, GREEN]
    for i, (tv, col) in enumerate(zip(thresholds, cols)):
        tv2 = tv*100 if fmt=="pct" else tv
        steps.append(dict(range=[prev, tv2], color=col+"22"))
        prev = tv2
    steps.append(dict(range=[prev, vmax], color=GREEN+"22"))
    fig = go.Figure(go.Indicator(mode="gauge+number",
        value=display,
        title=dict(text=title, font=dict(size=11, color=NAVY)),
        number=dict(suffix="%" if fmt=="pct" else "x", font=dict(size=18, color=NAVY)),
        gauge=dict(axis=dict(range=[vmin, vmax], tickfont=dict(size=8)),
            bar=dict(color=TEAL, thickness=0.35),
            steps=steps, bgcolor="white",
            borderwidth=1, bordercolor="#E2E8F0")))
    fig.update_layout(height=200, margin=dict(l=20,r=20,t=40,b=10),
        paper_bgcolor="white")
    return fig

def ratio_table(rows):
    html = '<table style="width:100%;border-collapse:collapse;font-size:12px">'
    html += '<tr style="background:#1F497D;color:white">'
    for h in ["Ratio","FY2025","FY2024","FY2023","FY2022","FY2021","Signal"]:
        html += f'<th style="padding:6px 8px;text-align:left;font-weight:600">{h}</th>'
    html += "</tr>"
    for i, row in enumerate(rows):
        bg = "#F8FAFB" if i%2==0 else "white"
        sig, sc = row[-1]
        cols_map = {GREEN:"#D4EDDA", AMBER:"#FFF3CD", RED:"#F8D7DA"}
        sbg = cols_map.get(sc, "#E8ECF0")
        html += f'<tr style="background:{bg}">'
        for cell in row[:-1]:
            html += f'<td style="padding:5px 8px;border-bottom:1px solid #F0F0F0">{cell}</td>'
        html += f'<td style="padding:5px 8px;border-bottom:1px solid #F0F0F0"><span style="background:{sbg};color:{sc};padding:2px 8px;border-radius:4px;font-weight:600;font-size:11px">{sig}</span></td>'
        html += "</tr>"
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

def sig(val, low_g, low_a, inverse=False):
    try:
        v = float(val)
        if not inverse:
            if v >= low_g: return ("Strong", GREEN)
            if v >= low_a: return ("Watch", AMBER)
            return ("Risk", RED)
        else:
            if v <= low_g: return ("Low", GREEN)
            if v <= low_a: return ("Moderate", AMBER)
            return ("High", RED)
    except: return ("--", DGRAY)

# ═══════════════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="background:linear-gradient(135deg,{NAVY} 0%,{TEAL} 100%);padding:20px 28px;border-radius:12px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center">
  <div>
    <div style="color:white;font-size:22px;font-weight:700;letter-spacing:.02em">⚡ Schneider Electric</div>
    <div style="color:rgba(255,255,255,.75);font-size:13px;margin-top:3px">Financial Dashboard · FY2021–FY2025 · Consolidated Annual Reports · Figures in EURm</div>
  </div>
  <div style="text-align:right">
    <div style="background:#1A7A3C;color:white;padding:6px 18px;border-radius:20px;font-weight:700;font-size:14px">STRONG</div>
    <div style="color:rgba(255,255,255,.7);font-size:11px;margin-top:4px">EBIT Margin {safe(d['ebit_m'][4],'pct')} · ROE {safe(d['roe'][4],'pct')}</div>
  </div>
</div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════
tabs = st.tabs(["Overview", "Profitability", "Liquidity", "Leverage", "Cash Flow", "Diagnosis"])

# ─── TAB 1: OVERVIEW ───────────────────────────────────────────────
with tabs[0]:
    section("Key Performance Indicators")
    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card("Revenue FY2025 (EURm)", safe(d['revenue'][4],'eur'), delta_abs(d['revenue'][4],d['revenue'][3]))
    with c2: kpi_card("EBIT Margin", safe(d['ebit_m'][4],'pct'), delta_str(d['ebit_m'][4],d['ebit_m'][3]))
    with c3: kpi_card("Free Cash Flow (EURm)", safe(d['fcf'][4],'eur'), delta_abs(d['fcf'][4],d['fcf'][3]))
    with c4: kpi_card("FCF Margin", safe(d['fcf_m'][4],'pct'), delta_str(d['fcf_m'][4],d['fcf_m'][3]), badge="NEW")

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card("ROE", safe(d['roe'][4],'pct'), delta_str(d['roe'][4],d['roe'][3]))
    with c2: kpi_card("Net Debt / EBITDA", safe(d['nd_ebitda'][4],'x'), delta_str(d['nd_ebitda'][4],d['nd_ebitda'][3], inverse=True), badge="NEW")
    with c3: kpi_card("Current Ratio", safe(d['current_r'][4],'x'), delta_str(d['current_r'][4],d['current_r'][3]))
    with c4: kpi_card("ROCE", safe(d['roce'][4],'pct'), delta_str(d['roce'][4],d['roce'][3]))

    section("Trend Analysis")
    c1,c2,c3 = st.columns(3)
    with c1:
        st.plotly_chart(line_chart("Margin Evolution (%)",
            {"Gross Margin": d['gross_m'], "EBIT Margin": d['ebit_m'], "Net Margin": d['ros']}),
            use_container_width=True)
    with c2:
        st.plotly_chart(bar_chart("Revenue & Gross Profit (EURm)",
            {"Revenue": d['revenue'], "Gross Profit": d['gross_profit']}),
            use_container_width=True)
    with c3:
        st.plotly_chart(bar_chart("Cash Flow Breakdown (EURm)",
            {"CFO": d['cfo'], "CFI": d['cfi'], "CFF": d['cff']}),
            use_container_width=True)

    section("Asset & Funding Structure (FY2025)")
    c1,c2,c3 = st.columns(3)
    with c1:
        st.plotly_chart(donut(
            ["Intangibles","PP&E","Other NCA","Inventories","Receivables","Cash"],
            [d['intangibles'][0],d['ppe'][0],d['oth_nca'][0],d['inventories'][0],d['receivables'][0],d['cash'][0]],
            "Asset Mix", [NAVY,TEAL,GREEN,AMBER,RED,"#8E44AD"]), use_container_width=True)
    with c2:
        st.plotly_chart(donut(
            ["Equity","LT Debt","ST Debt"],
            [d['equity'][0],d['lt_debt'][0],d['st_debt'][0]],
            "Funding Structure", [GREEN,NAVY,TEAL]), use_container_width=True)
    with c3:
        capex_abs = abs(d['capex'][4])
        divs_abs  = abs(d['divs'][4])
        buy_abs   = abs(d['buyback'][4])
        ma_abs    = abs(d['ma'][4])
        st.plotly_chart(donut(
            ["Dividends","CapEx","Buybacks","M&A"],
            [divs_abs, capex_abs, buy_abs, ma_abs],
            "Capital Allocation FY2025", [TEAL,NAVY,GREEN,AMBER]), use_container_width=True)

# ─── TAB 2: PROFITABILITY ──────────────────────────────────────────
with tabs[1]:
    section("Profit Margins — All Years")
    ratio_table([
        ["Gross Margin",     *[safe(d['gross_m'][i],'pct') for i in [4,3,2,1,0]],  sig(d['gross_m'][4], 0.38, 0.30)],
        ["EBITA Margin",     *[safe(d['ebita_m'][i],'pct') for i in [4,3,2,1,0]],  sig(d['ebita_m'][4], 0.16, 0.10)],
        ["EBIT Margin",      *[safe(d['ebit_m'][i],'pct') for i in [4,3,2,1,0]],   sig(d['ebit_m'][4], 0.15, 0.10)],
        ["EBITDA Margin",    *[safe(d['ebitda_m'][i],'pct') for i in [4,3,2,1,0]], sig(d['ebitda_m'][4], 0.18, 0.12)],
        ["Net Margin (ROS)", *[safe(d['ros'][i],'pct') for i in [4,3,2,1,0]],      sig(d['ros'][4], 0.10, 0.06)],
    ])

    c1, c2 = st.columns(2)
    with c1:
        section("Margin Evolution")
        st.plotly_chart(line_chart("",
            {"Gross":d['gross_m'],"EBIT":d['ebit_m'],"EBITDA":d['ebitda_m'],"Net":d['ros']}),
            use_container_width=True)
    with c2:
        section("Scissor Effect — Revenue vs COGS Growth")
        rev_g  = [None]+[(d['revenue'][i]-d['revenue'][i-1])/d['revenue'][i-1] for i in range(1,5)]
        cogs_g = [None]+[(abs(d['cogs'][i])-abs(d['cogs'][i-1]))/abs(d['cogs'][i-1]) for i in range(1,5)]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=YEARS, y=[v*100 if v else None for v in rev_g],
            name="Revenue Growth", line=dict(color=GREEN,width=2.5), mode="lines+markers"))
        fig.add_trace(go.Scatter(x=YEARS, y=[v*100 if v else None for v in cogs_g],
            name="COGS Growth", line=dict(color=RED,width=2.5,dash="dash"), mode="lines+markers"))
        fig.update_layout(height=280,margin=dict(l=0,r=0,t=10,b=0),
            plot_bgcolor="white",paper_bgcolor="white",
            yaxis=dict(ticksuffix="%",gridcolor="#F0F0F0"),
            legend=dict(orientation="h",y=-0.2,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)

    section("DuPont Decomposition — ROE = ROS x Asset Turnover x Equity Multiplier")
    ratio_table([
        ["ROE",               *[safe(d['roe'][i],'pct') for i in [4,3,2,1,0]],      sig(d['roe'][4], 0.15, 0.08)],
        ["ROS (Net Margin)",  *[safe(d['ros'][i],'pct') for i in [4,3,2,1,0]],      sig(d['ros'][4], 0.10, 0.06)],
        ["Asset Turnover",    *[safe(v,'x')   for v in reversed(d['asset_to'])], sig(d['asset_to'][4], 0.6, 0.4)],
        ["Equity Multiplier", *[safe(v,'x')   for v in reversed(d['eq_mult'])],  sig(d['eq_mult'][4], 0, 0, inverse=False)],
        ["ROA",               *[safe(d['ros'][i]*d['asset_to'][i],'pct') for i in [4,3,2,1,0]], sig(d['ros'][4]*d['asset_to'][4], 0.06, 0.04)],
        ["ROCE",              *[safe(d['roce'][i],'pct') for i in [4,3,2,1,0]],      sig(d['roce'][4], 0.10, 0.07)],
    ])

    section("Revenue & Profit Waterfall (EURm)")
    st.plotly_chart(bar_chart("",
        {"Revenue":d['revenue'],"Gross Profit":d['gross_profit'],
         "EBIT":d['ebit'],"Net Income":d['net_income']}),
        use_container_width=True)

# ─── TAB 3: LIQUIDITY ──────────────────────────────────────────────
with tabs[2]:
    section("Liquidity Ratios")
    c1,c2,c3 = st.columns(3)
    with c1: st.plotly_chart(gauge(d['current_r'][4],"Current Ratio",0,2.5,[1.0,1.5],"x"), use_container_width=True)
    with c2: st.plotly_chart(gauge(d['quick_r'][4],"Quick Ratio",0,2,[0.7,1.0],"x"), use_container_width=True)
    with c3: st.plotly_chart(gauge(d['cash_r'][4],"Cash Ratio",0,1,[0.2,0.4],"x"), use_container_width=True)

    ratio_table([
        ["Current Ratio", *[safe(d['current_r'][i],'x') for i in [4,3,2,1,0]], sig(d['current_r'][4],1.5,1.0)],
        ["Quick Ratio",   *[safe(d['quick_r'][i],'x') for i in [4,3,2,1,0]],  sig(d['quick_r'][4],1.0,0.7)],
        ["Cash Ratio",    *[safe(d['cash_r'][i],'x') for i in [4,3,2,1,0]],   sig(d['cash_r'][4],0.4,0.2)],
    ])

    section("Working Capital Cycle")
    c1,c2 = st.columns([3,2])
    with c1:
        ratio_table([
            ["Days Sales Outstanding (DSO)",   *[safe(d['dso'][i],'days') for i in [4,3,2,1,0]],  sig(d['dso'][4],0, 90, inverse=True)],
            ["Days Inventory Outstanding (DIO)",*[safe(d['dio'][i],'days') for i in [4,3,2,1,0]], sig(d['dio'][4],0, 90, inverse=True)],
            ["Days Payable Outstanding (DPO)",  *[safe(d['dpo'][i],'days') for i in [4,3,2,1,0]], sig(d['dpo'][4],90,60)],
            ["Cash Conversion Cycle (CCC)",     *[safe(d['ccc'][i],'days') for i in [4,3,2,1,0]], sig(d['ccc'][4],0, 40, inverse=True)],
        ])
    with c2:
        st.plotly_chart(line_chart("CCC Evolution (days)",
            {"DSO":d['dso'],"DIO":d['dio'],"CCC":d['ccc']}, y_fmt="raw", height=220),
            use_container_width=True)

    section("Liquidity Trend")
    st.plotly_chart(line_chart("Current & Quick Ratio Evolution",
        {"Current Ratio":d['current_r'],"Quick Ratio":d['quick_r'],"Cash Ratio":d['cash_r']},
        y_fmt="raw", height=250), use_container_width=True)

# ─── TAB 4: LEVERAGE ───────────────────────────────────────────────
with tabs[3]:
    section("Solvency Ratios — EBITA & EBITDA Versions")
    ratio_table([
        ["Debt / Equity",            *[safe(v,'x')   for v in reversed(d['de_ratio'])],   sig(d['de_ratio'][4],0,1.0,inverse=True)],
        ["Total Debt Ratio",         *[safe(d['debt_ratio'][i],'pct') for i in [4,3,2,1,0]], sig(d['debt_ratio'][4],0,0.35,inverse=True)],
        ["Interest Coverage (EBIT)", *[safe(v,'x')   for v in reversed(d['int_cov'])],    sig(d['int_cov'][4],5,3)],
        ["Interest Cov. (EBITDA)",   *[safe(v,'x')   for v in reversed(d['int_cov_ebd'])],sig(d['int_cov_ebd'][4],6,4)],
        ["Net Debt / EBITA",         *[safe(v,'x')   for v in reversed(d['nd_ebita'])],   sig(d['nd_ebita'][4],0,2.5,inverse=True)],
        ["Net Debt / EBITDA",        *[safe(v,'x')   for v in reversed(d['nd_ebitda'])],  sig(d['nd_ebitda'][4],0,3.0,inverse=True)],
        ["ROCE",                     *[safe(d['roce'][i],'pct') for i in [4,3,2,1,0]],        sig(d['roce'][4],0.10,0.07)],
    ])

    c1,c2 = st.columns(2)
    with c1:
        section("Net Debt Evolution (EURm)")
        fig = go.Figure()
        nd_colors = [GREEN if v<10000 else AMBER if v<13000 else RED for v in d['net_debt']]
        fig.add_trace(go.Bar(x=YEARS, y=d['net_debt'], marker_color=nd_colors,
            text=[f"{v:,.0f}" for v in d['net_debt']], textposition="outside", textfont=dict(size=10)))
        fig.update_layout(height=280,margin=dict(l=0,r=0,t=10,b=0),
            plot_bgcolor="white",paper_bgcolor="white",
            yaxis=dict(gridcolor="#F0F0F0"),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        section("Coverage Ratios")
        st.plotly_chart(line_chart("",
            {"Int. Coverage EBIT":d['int_cov'],"Int. Coverage EBITDA":d['int_cov_ebd']},
            y_fmt="raw", height=280), use_container_width=True)

    section("Capital Structure FY2025")
    c1,c2 = st.columns(2)
    with c1:
        st.plotly_chart(donut(
            ["Equity","LT Debt","ST Debt"],
            [d['equity'][0],d['lt_debt'][0],d['st_debt'][0]],
            "Debt vs Equity", [GREEN,NAVY,TEAL]), use_container_width=True)
    with c2:
        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:20px;border:1px solid #E2E8F0;margin-top:10px">
          <div style="font-size:12px;color:{DGRAY};margin-bottom:16px;font-weight:600">CAPITAL STRUCTURE SUMMARY</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
            <div><div style="font-size:10px;color:{DGRAY}">Total Equity</div><div style="font-size:18px;font-weight:700;color:{GREEN}">{safe(d['equity'][0],'eur')} €m</div></div>
            <div><div style="font-size:10px;color:{DGRAY}">LT Debt</div><div style="font-size:18px;font-weight:700;color:{NAVY}">{safe(d['lt_debt'][0],'eur')} €m</div></div>
            <div><div style="font-size:10px;color:{DGRAY}">ST Debt</div><div style="font-size:18px;font-weight:700;color:{TEAL}">{safe(d['st_debt'][0],'eur')} €m</div></div>
            <div><div style="font-size:10px;color:{DGRAY}">Net Debt</div><div style="font-size:18px;font-weight:700;color:{RED}">{safe(d['net_debt'][4],'eur')} €m</div></div>
          </div>
        </div>""", unsafe_allow_html=True)

# ─── TAB 5: CASH FLOW ──────────────────────────────────────────────
with tabs[4]:
    section("Cash Flow Summary (EURm)")
    ratio_table([
        ["Cash from Operations (CFO)", *[safe(d['cfo'][i],'eur') for i in [4,3,2,1,0]], sig(d['cfo'][4],5000,3000)],
        ["Cash from Investing (CFI)",  *[safe(d['cfi'][i],'eur') for i in [4,3,2,1,0]], ("Negative by nature", DGRAY)],
        ["Cash from Financing (CFF)",  *[safe(d['cff'][i],'eur') for i in [4,3,2,1,0]], ("Negative by nature", DGRAY)],
        ["Net CapEx",                  *[safe(d['capex'][i],'eur') for i in [4,3,2,1,0]],("Investing", DGRAY)],
        ["Free Cash Flow (FCF)",       *[safe(d['fcf'][i],'eur') for i in [4,3,2,1,0]],  sig(d['fcf'][4],4000,2000)],
    ])

    section("FCF Quality Ratios")
    ratio_table([
        ["FCF Margin",       *[safe(d['fcf_m'][i],'pct') for i in [4,3,2,1,0]],   sig(d['fcf_m'][4],0.10,0.06)],
        ["FCF / Net Income", *[safe(v,'x')   for v in reversed(d['fcf_ni'])],  sig(d['fcf_ni'][4],1.0,0.7)],
        ["CFO / EBITDA",     *[safe(v,'x')   for v in reversed(d['cfo_ebd'])], sig(d['cfo_ebd'][4],0.75,0.55)],
        ["CapEx / Revenue",  *[safe(d['capex_r'][i],'pct') for i in [4,3,2,1,0]], ("Reinvestment", DGRAY)],
        ["CapEx / D&A",      *[safe(v,'x')   for v in reversed(d['capex_da'])],("Maintenance", DGRAY)],
    ])

    c1,c2 = st.columns(2)
    with c1:
        section("CFO vs FCF Evolution")
        st.plotly_chart(bar_chart("",{"CFO":d['cfo'],"FCF":d['fcf']}), use_container_width=True)
    with c2:
        section("Capital Allocation FY2025 (EURm)")
        capex_a = abs(d['capex'][4])
        divs_a  = abs(d['divs'][4])
        buy_a   = abs(d['buyback'][4])
        ma_a    = abs(d['ma'][4])
        st.plotly_chart(donut(
            [f"Dividends ({divs_a:,.0f})",f"CapEx ({capex_a:,.0f})",
             f"Buybacks ({buy_a:,.0f})",f"M&A ({ma_a:,.0f})"],
            [divs_a,capex_a,buy_a,ma_a],
            "", [TEAL,NAVY,GREEN,AMBER]), use_container_width=True)

# ─── TAB 6: DIAGNOSIS ──────────────────────────────────────────────
with tabs[5]:
    ebit_m25 = d['ebit_m'][4]
    verdict = "STRONG" if ebit_m25 > 0.15 else "STABLE" if ebit_m25 >= 0.10 else "CONCERNING"
    v_color = GREEN if verdict=="STRONG" else AMBER if verdict=="STABLE" else RED

    st.markdown(f"""
    <div style="background:white;border-radius:12px;padding:20px 24px;border:1px solid #E2E8F0;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center">
      <div>
        <div style="font-size:11px;font-weight:600;color:{DGRAY};text-transform:uppercase;letter-spacing:.06em">Overall Financial Health</div>
        <div style="font-size:32px;font-weight:800;color:{v_color};margin-top:2px">{verdict}</div>
        <div style="font-size:12px;color:{DGRAY};margin-top:4px">Source: Consolidated Annual Report FY2025 · Figures in EURm</div>
      </div>
      <div style="text-align:right;font-size:12px;color:{DGRAY}">
        <div>EBIT Margin: <b style="color:{NAVY}">{safe(d['ebit_m'][4],'pct')}</b></div>
        <div>ROE: <b style="color:{NAVY}">{safe(d['roe'][4],'pct')}</b></div>
        <div>FCF: <b style="color:{NAVY}">{safe(d['fcf'][4],'eur')} EURm</b></div>
        <div>Net Debt/EBITDA: <b style="color:{NAVY}">{safe(d['nd_ebitda'][4],'x')}</b></div>
      </div>
    </div>""", unsafe_allow_html=True)

    section("Pillar Summary")
    pillar_data = [
        ["Liquidity",       "Current Ratio",      safe(d['current_r'][4],'x'),   safe(d['current_r'][3],'x'),   sig(d['current_r'][4],1.5,1.0)],
        ["Profitability",   "EBIT Margin",         safe(d['ebit_m'][4],'pct'),    safe(d['ebit_m'][3],'pct'),    sig(d['ebit_m'][4],0.15,0.10)],
        ["Leverage",        "Net Debt/EBITDA",     safe(d['nd_ebitda'][4],'x'),   safe(d['nd_ebitda'][3],'x'),   sig(d['nd_ebitda'][4],0,2.5,inverse=True)],
        ["Efficiency",      "CCC (days)",          safe(d['ccc'][4],'days'),      safe(d['ccc'][3],'days'),      sig(d['ccc'][4],0,40,inverse=True)],
        ["Cash Generation", "FCF Margin",          safe(d['fcf_m'][4],'pct'),     safe(d['fcf_m'][3],'pct'),     sig(d['fcf_m'][4],0.10,0.06)],
        ["Investors",       "ROE",                 safe(d['roe'][4],'pct'),       safe(d['roe'][3],'pct'),       sig(d['roe'][4],0.15,0.10)],
    ]
    html = '<table style="width:100%;border-collapse:collapse;font-size:12px">'
    html += f'<tr style="background:{NAVY};color:white"><th style="padding:7px 10px;text-align:left">Pillar</th><th>Key Ratio</th><th>FY2025</th><th>FY2024</th><th>Signal</th></tr>'
    for i,row in enumerate(pillar_data):
        bg="#F8FAFB" if i%2==0 else "white"
        sig_txt, sc = row[4]
        cols_map={GREEN:"#D4EDDA",AMBER:"#FFF3CD",RED:"#F8D7DA"}
        sbg=cols_map.get(sc,"#E8ECF0")
        html+=f'<tr style="background:{bg}"><td style="padding:7px 10px;font-weight:600;color:{NAVY};border-bottom:1px solid #F0F0F0">{row[0]}</td>'
        html+=f'<td style="padding:7px 10px;border-bottom:1px solid #F0F0F0">{row[1]}</td>'
        html+=f'<td style="padding:7px 10px;font-weight:700;border-bottom:1px solid #F0F0F0">{row[2]}</td>'
        html+=f'<td style="padding:7px 10px;color:{DGRAY};border-bottom:1px solid #F0F0F0">{row[3]}</td>'
        html+=f'<td style="padding:7px 10px;border-bottom:1px solid #F0F0F0"><span style="background:{sbg};color:{sc};padding:2px 10px;border-radius:4px;font-weight:600;font-size:11px">{sig_txt}</span></td></tr>'
    html+="</table>"
    st.markdown(html, unsafe_allow_html=True)

    section("Written Analysis (Complete these cells)")
    c1,c2 = st.columns(2)
    def yellow_box(title):
        st.markdown(f"""<div style="background:#FFFDE7;border:1px solid #F9A825;border-radius:8px;padding:14px;min-height:100px">
        <div style="font-size:11px;font-weight:700;color:#F57F17;margin-bottom:8px;text-transform:uppercase">{title}</div>
        <div style="font-size:11px;color:#888;font-style:italic">Write your analysis here...</div>
        </div>""", unsafe_allow_html=True)
    with c1:
        yellow_box("Liquidity Analysis")
        st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)
        yellow_box("Leverage Analysis")
    with c2:
        yellow_box("Profitability Analysis")
        st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)
        yellow_box("Efficiency Analysis")

    section("Strengths & Red Flags")
    c1,c2 = st.columns(2)
    with c1:
        st.markdown(f"""<div style="background:#F0FFF4;border:1px solid #9AE6B4;border-radius:8px;padding:14px">
        <div style="font-size:11px;font-weight:700;color:{GREEN};margin-bottom:10px">KEY STRENGTHS</div>
        {"".join(f'<div style="font-size:12px;color:#2D3748;margin-bottom:6px;padding:4px 8px;background:white;border-radius:4px">+ Write strength {i} here</div>' for i in range(1,4))}
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div style="background:#FFF5F5;border:1px solid #FEB2B2;border-radius:8px;padding:14px">
        <div style="font-size:11px;font-weight:700;color:{RED};margin-bottom:10px">RED FLAGS / RISKS</div>
        {"".join(f'<div style="font-size:12px;color:#2D3748;margin-bottom:6px;padding:4px 8px;background:white;border-radius:4px">- Write risk {i} here</div>' for i in range(1,4))}
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>",unsafe_allow_html=True)
    st.markdown(f"""<div style="background:#FFFDE7;border:1px solid #F9A825;border-radius:8px;padding:14px">
    <div style="font-size:11px;font-weight:700;color:#F57F17;margin-bottom:8px;text-transform:uppercase">Recommendation</div>
    <div style="font-size:11px;color:#888;font-style:italic">Write your overall recommendation here...</div>
    </div>""", unsafe_allow_html=True)
