import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd

st.set_page_config(page_title="Schneider Electric - Financial Dashboard", layout="wide")

# ── THEME ──────────────────────────────────────────────────────────
NAVY   = "#111111"
TEAL   = "#166534"
GREEN  = "#16A34A"
AMBER  = "#6B7280"
RED    = "#C0392B"
LGRAY  = "#F5F7FA"
MGRAY  = "#E8ECF0"
DGRAY  = "#4B5563"
WHITE  = "#FFFFFF"
YEARS  = [2021, 2022, 2023, 2024, 2025]

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Aptos Narrow', 'Aptos', 'Arial Narrow', Arial, sans-serif; }
.main { background: #F8FAFB; }
.block-container { padding: 1.2rem 2rem 2rem; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #E8ECF0; border-radius: 10px; padding: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 18px; font-weight: 500; font-size: 13px; color: #111111 !important; }
.stTabs [aria-selected="true"] { background: #166534 !important; color: white !important; }
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

    # IS cols ordered FY2021→FY2025: 28,25,22,19,16
    IC = [28, 25, 22, 19, 16]
    # Ratios cols ordered FY2021→FY2025: 9,8,7,6,5
    RC = [9, 8, 7, 6, 5]
    # CFS cols ordered FY2021→FY2025: 26,24,22,20,18
    CC = [26, 24, 22, 20, 18]

    def irow(df, c): return [df.iloc[c, x] for x in IC]
    def rrow(df, c): return [df.iloc[c, x] for x in RC]
    def crow(df, c): return [df.iloc[c, x] for x in CC]

    d = {}

    # Income Statement (0-based; row 2 = headers, row 3 = first data row)
    d["revenue"]      = irow(i, 3)
    d["cogs"]         = irow(i, 4)
    d["gross_profit"] = irow(i, 5)
    d["ebita"]        = irow(i, 11)
    d["ebitda"]       = irow(i, 12)
    d["ebit"]         = irow(i, 14)
    d["net_income"]   = irow(i, 24)
    d["int_expense"]  = irow(i, 16)  # gross cost of financial debt

    # Ratios (0-based; row 2 = headers, row 3 = first data row)
    d["roe"]          = rrow(r, 3)
    d["ros"]          = rrow(r, 4)
    d["asset_to"]     = rrow(r, 5)
    d["eq_mult"]      = rrow(r, 6)
    d["gross_m"]      = rrow(r, 10)
    d["ebita_m"]      = rrow(r, 13)
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
    d["payout"]       = rrow(r, 34)
    d["div_yield"]    = rrow(r, 35)
    d["eps"]          = rrow(r, 36)
    d["pe"]           = rrow(r, 37)
    d["pb"]           = rrow(r, 38)
    d["net_debt"]     = rrow(r, 39)
    d["ev"]           = rrow(r, 40)
    d["ev_ebita"]     = rrow(r, 41)
    d["roce"]         = rrow(r, 43)

    # Derived ratios
    d["ebitda_m"]     = [e/rev if rev else None for e,rev in zip(d["ebitda"], d["revenue"])]
    d["nd_ebitda"]    = [nd/eb if eb else None for nd,eb in zip(d["net_debt"], d["ebitda"])]
    d["int_cov_ebd"]  = [eb/abs(ie) if ie else None for eb,ie in zip(d["ebitda"], d["int_expense"])]

    # Cash Flow Statement (0-based; row 0 = headers, row 1 = first data row)
    d["cfo"]     = crow(cf, 19)   # TOTAL I – operating activities
    d["capex"]   = crow(cf, 23)   # Net capital expenditure
    d["cfi"]     = crow(cf, 28)   # TOTAL II – investing activities
    d["cff"]     = crow(cf, 38)   # TOTAL III – financing activities
    d["divs"]    = crow(cf, 36)   # Dividends paid to SE shareholders
    d["buyback"] = crow(cf, 31)   # Purchase of treasury shares
    d["ma"]      = crow(cf, 24)   # Acquisitions and disposals, net

    d["fcf"]     = [c + cap for c,cap in zip(d["cfo"], d["capex"])]
    d["fcf_m"]   = [f/rev if rev else None for f,rev in zip(d["fcf"], d["revenue"])]
    d["fcf_ni"]  = [f/ni if ni else None for f,ni in zip(d["fcf"], d["net_income"])]
    d["cfo_ebd"] = [c/e if e else None for c,e in zip(d["cfo"], d["ebitda"])]
    d["capex_r"] = [abs(cap)/rev if rev else None for cap,rev in zip(d["capex"], d["revenue"])]
    dda = [e - eb for e,eb in zip(d["ebitda"], d["ebita"])]
    d["capex_da"]= [abs(cap)/da if da else None for cap,da in zip(d["capex"], dda)]

    # Balance Sheet (0-based; row 0 = headers, col 15 = FY2025, col 18 = FY2024)
    d["intangibles"]  = [bs.iloc[2,  15], bs.iloc[2,  18]]
    d["ppe"]          = [bs.iloc[3,  15], bs.iloc[3,  18]]
    d["oth_nca"]      = [bs.iloc[4,  15], bs.iloc[4,  18]]
    d["inventories"]  = [bs.iloc[6,  15], bs.iloc[6,  18]]
    d["receivables"]  = [bs.iloc[7,  15], bs.iloc[7,  18]]
    d["cash"]         = [bs.iloc[9,  15], bs.iloc[9,  18]]
    d["total_assets"] = [bs.iloc[11, 15], bs.iloc[11, 18]]
    d["equity"]       = [bs.iloc[16, 15], bs.iloc[16, 18]]
    d["lt_debt"]      = [bs.iloc[17, 15], bs.iloc[17, 18]]
    d["st_debt"]      = [bs.iloc[22, 15], bs.iloc[22, 18]]

    return d

_d_base = load()

# ── MANUAL ENTRY MERGE ─────────────────────────────────────────────
# BS keys are 2-element snapshots, not year-indexed time series
_BS_KEYS = {'intangibles','ppe','oth_nca','inventories','receivables',
            'cash','total_assets','equity','lt_debt','st_debt'}
d = {k: list(v) for k, v in _d_base.items()}

if 'manual_entries' not in st.session_state:
    st.session_state.manual_entries = {}

for _yr in sorted(st.session_state.manual_entries.keys()):
    YEARS.append(_yr)
    _e = st.session_state.manual_entries[_yr]
    for _k in d:
        if _k not in _BS_KEYS:
            d[_k].append(_e.get(_k))

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
        pairs = [(y, v*100 if y_fmt=="pct" else v) for y, v in zip(YEARS, vals) if v is not None]
        xs = [p[0] for p in pairs]
        yv = [p[1] for p in pairs]
        fig.add_trace(go.Scatter(x=xs, y=yv, name=name,
            line=dict(color=colors[i % len(colors)], width=2.5),
            mode="lines+markers", marker=dict(size=6)))
    fig.update_layout(height=height, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        font=dict(color="#111111"),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.2, font=dict(size=10, color="#111111")),
        yaxis=dict(ticksuffix="%" if y_fmt=="pct" else "", gridcolor="#F0F0F0",
                   tickfont=dict(size=10, color="#111111")),
        xaxis=dict(tickfont=dict(size=10, color="#111111"), dtick=1))
    return fig

def bar_chart(title, series_dict, height=280, stack=False):
    fig = go.Figure()
    colors = [TEAL, NAVY, GREEN, AMBER, RED, "#6B7280"]
    bt = "stack" if stack else "group"
    for i, (name, vals) in enumerate(series_dict.items()):
        c = colors[i % len(colors)]
        clean = [v if v is not None else 0 for v in vals]
        fig.add_trace(go.Bar(x=YEARS, y=clean, name=name,
            marker_color=c, opacity=0.88))
    fig.update_layout(barmode=bt, height=height, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        font=dict(color="#111111"),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.2, font=dict(size=10, color="#111111")),
        yaxis=dict(gridcolor="#F0F0F0", tickfont=dict(size=10, color="#111111")),
        xaxis=dict(tickfont=dict(size=10, color="#111111"), dtick=1))
    return fig

def donut(labels, values, title, colors_list):
    fig = go.Figure(go.Pie(labels=labels, values=[abs(v) for v in values],
        hole=0.62, marker_colors=colors_list,
        textinfo="percent", textfont=dict(size=10, color="white"),
        hovertemplate="%{label}: %{value:,.0f} EURm<extra></extra>"))
    fig.update_layout(height=220, margin=dict(l=0,r=0,t=30,b=0),
        title=dict(text=title, font=dict(size=12, color=NAVY), x=0),
        font=dict(color="#111111"),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=True, legend=dict(font=dict(size=9, color="#111111"), orientation="h", y=-0.1))
    return fig

def gauge(value, title, min_v, max_v, thresholds, fmt="pct"):
    display = value*100 if fmt=="pct" else value
    vmin = min_v*100 if fmt=="pct" else min_v
    vmax = max_v*100 if fmt=="pct" else max_v
    steps = []
    prev = vmin
    zone_colors = ["rgba(192,57,43,0.12)", "rgba(196,120,0,0.12)",
                   "rgba(26,122,60,0.12)"]
    for idx, (tv, col) in enumerate(zip(thresholds, zone_colors)):
        tv2 = tv*100 if fmt=="pct" else tv
        steps.append(dict(range=[prev, tv2], color=col))
        prev = tv2
    steps.append(dict(range=[prev, vmax], color="rgba(26,122,60,0.12)"))
    fig = go.Figure(go.Indicator(mode="gauge+number",
        value=display,
        title=dict(text=title, font=dict(size=11, color="#111111")),
        number=dict(suffix="%" if fmt=="pct" else "x", font=dict(size=18, color="#111111")),
        gauge=dict(axis=dict(range=[vmin, vmax], tickfont=dict(size=8, color="#111111")),
            bar=dict(color=TEAL, thickness=0.35),
            steps=steps, bgcolor="white",
            borderwidth=1, bordercolor="#E2E8F0")))
    fig.update_layout(height=200, margin=dict(l=20,r=20,t=40,b=10),
        font=dict(color="#111111"),
        paper_bgcolor="white")
    return fig

def ratio_table(rows):
    html = '<table style="width:100%;border-collapse:collapse;font-size:12px">'
    html += '<tr style="background:#111111">'
    for h in ["Ratio", "FY2025", "FY2024", "FY2023", "FY2022", "FY2021", "Signal"]:
        html += f'<th style="padding:7px 10px;text-align:left;font-weight:600;color:white">{h}</th>'
    html += "</tr>"
    for i, row in enumerate(rows):
        bg = "#F8FAFB" if i % 2 == 0 else "white"
        sig_text, sc = row[-1]
        badge_bg = {GREEN: "#D4EDDA", AMBER: "#F3F4F6", RED: "#F8D7DA"}.get(sc, "#F3F4F6")
        html += f'<tr style="background:{bg}">'
        for cell in row[:-1]:
            html += f'<td style="padding:6px 10px;border-bottom:1px solid #F0F0F0;color:#111111">{cell}</td>'
        html += f'<td style="padding:6px 10px;border-bottom:1px solid #F0F0F0"><span style="background:{badge_bg};color:{sc};padding:2px 8px;border-radius:4px;font-weight:600;font-size:11px">{sig_text}</span></td>'
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
    <div style="color:white;font-size:22px;font-weight:700;letter-spacing:.02em">Schneider Electric</div>
    <div style="color:rgba(255,255,255,.75);font-size:13px;margin-top:3px">Financial Dashboard · FY2021–FY2025 · Consolidated Annual Reports · Figures in EURm</div>
  </div>
  <div style="text-align:right">
    <div style="background:{GREEN};color:white;padding:6px 18px;border-radius:20px;font-weight:700;font-size:14px">STRONG</div>
    <div style="color:rgba(255,255,255,.7);font-size:11px;margin-top:4px">EBIT Margin {safe(d['ebit_m'][4],'pct')} · ROE {safe(d['roe'][4],'pct')}</div>
  </div>
</div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════
tabs = st.tabs(["Overview", "Profitability", "Liquidity", "Leverage", "Cash Flow", "Diagnosis", "Data Entry"])

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
            font=dict(color="#111111"),
            yaxis=dict(ticksuffix="%",gridcolor="#F0F0F0",tickfont=dict(size=10,color="#111111")),
            xaxis=dict(tickmode="array", tickvals=[2021,2022,2023,2024,2025],
                       ticktext=["2021","2022","2023","2024","2025"],
                       tickfont=dict(size=10,color="#111111")),
            legend=dict(orientation="h",y=-0.2,font=dict(size=10,color="#111111")))
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
    section("Liquidity Ratios — FY2025 Snapshot")
    c1,c2,c3 = st.columns(3)
    with c1:
        st.plotly_chart(gauge(d['current_r'][4],"Current Ratio",0,2.5,[1.0,1.5],"x"), use_container_width=True)
        st.markdown(f"""<div style="background:#F8FAFB;border-radius:8px;padding:10px 14px;font-size:11px;color:#4B5563;border:1px solid #E2E8F0">
        <b>Formula:</b> Current Assets / Current Liabilities<br>
        <b>Benchmark:</b> &gt;1.5 Strong · 1.0–1.5 Watch · &lt;1.0 Risk<br>
        <b>YoY change:</b> {delta_str(d['current_r'][4], d['current_r'][3])}
        </div>""", unsafe_allow_html=True)
    with c2:
        st.plotly_chart(gauge(d['quick_r'][4],"Quick Ratio",0,2,[0.7,1.0],"x"), use_container_width=True)
        st.markdown(f"""<div style="background:#F8FAFB;border-radius:8px;padding:10px 14px;font-size:11px;color:#4B5563;border:1px solid #E2E8F0">
        <b>Formula:</b> (Current Assets − Inventories) / Current Liabilities<br>
        <b>Benchmark:</b> &gt;1.0 Strong · 0.7–1.0 Watch · &lt;0.7 Risk<br>
        <b>YoY change:</b> {delta_str(d['quick_r'][4], d['quick_r'][3])}
        </div>""", unsafe_allow_html=True)
    with c3:
        st.plotly_chart(gauge(d['cash_r'][4],"Cash Ratio",0,1,[0.2,0.4],"x"), use_container_width=True)
        st.markdown(f"""<div style="background:#F8FAFB;border-radius:8px;padding:10px 14px;font-size:11px;color:#4B5563;border:1px solid #E2E8F0">
        <b>Formula:</b> Cash & Equivalents / Current Liabilities<br>
        <b>Benchmark:</b> &gt;0.4 Strong · 0.2–0.4 Watch · &lt;0.2 Risk<br>
        <b>YoY change:</b> {delta_str(d['cash_r'][4], d['cash_r'][3])}
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    section("5-Year Liquidity Table")
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
            font=dict(color="#111111"),
            yaxis=dict(gridcolor="#F0F0F0",tickfont=dict(color="#111111")),
            xaxis=dict(tickfont=dict(color="#111111")),showlegend=False)
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
    html = '<table style="width:100%;border-collapse:collapse;font-size:12px;color:#111111">'
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

    section("Executive Summary")
    st.markdown("""<div style="background:#F0FFF4;border-left:4px solid #166534;border-radius:0 8px 8px 0;padding:18px 22px;margin-bottom:4px"><p style="font-size:13px;color:#1a202c;line-height:1.75;margin:0">Over the 2021–2025 period, Schneider Electric delivered consistent and profitable growth: revenue rose from €28.9bn to €40.2bn (+39%), the EBITA margin expanded from 16.4% to 17.8%, and net income grew from €3.3bn to €4.4bn. Cash generation is one of the group's strongest assets; operating cash flow reached €6.1bn in 2025, up nearly 70% over five years, and dividends were increased every single year. The balance sheet remains broadly healthy, though leverage increased significantly in 2025 following strategic acquisitions (AVEVA, SEIPL), with Net Debt/EBITDA reaching 1.86x, the highest level of the period and the main financial risk to monitor going forward. Two structural weaknesses also stand out: working capital needs grew more than three times faster than revenue (+127% vs +39%), and R&amp;D spending at ~3.4% of sales remains low for a group positioning itself as a technology company. Overall, Schneider is a financially solid, growing business, well-positioned on high-potential markets (electrification, data centres, AI infrastructure); the priority now is to reduce debt, step up innovation investment, and tighten working capital discipline to lock in long-term value creation.</p></div>""", unsafe_allow_html=True)

    section("Part I — Overall Financial Analysis")
    for _t, _b in [
        ("1. A MILESTONE YEAR: BREAKING THE €40 BILLION THRESHOLD",
         "FY2025 marks a historic milestone: Schneider Electric crossed the €40 billion revenue threshold for the first time, delivering +9% organic growth in a volatile macro environment. This performance is not circumstantial, it reflects five years of consistent execution across two structural megatrends: electrification and digitalization. Energy Management achieved double-digit organic growth for the fifth consecutive year, while Industrial Automation, which had struggled through a destocking cycle in 2024, returned to growth in 2025. The record backlog of over €25 billion, up +18% year-on-year, gives the group exceptional visibility heading into 2026 and beyond."),
        ("2. PROFITABILITY: ROBUST MARGINS DESPITE A COMPLEX ENVIRONMENT",
         "Schneider's profitability has steadily improved over the period. The adjusted EBITA margin expanded to 17.8% in 2025, up from 16.4% in 2021, supported by a gross margin consistently holding around 42%, a sign of genuine pricing power and a richer product mix. Return on Equity reached 17.8%, the highest level in five years. The one nuance worth flagging: net margin experienced a slight compression in 2025, driven not by operational weakness but by the cost of a deliberate leverage increase to fund strategic acquisitions. This is a trade-off the Board consciously accepted, and one that must now be carefully managed."),
        ("3. THE DIGITAL FLYWHEEL: SCHNEIDER'S STRATEGIC TRANSFORMATION ENGINE",
         "The Digital Flywheel — Schneider's integrated offer of connected products, software and services — is now the group's primary growth engine, generating €25 billion in revenue and representing 62% of total sales. The full acquisition of AVEVA in 2023 was the pivotal move underpinning this transformation, giving Schneider a world-class industrial software platform used by over 20,000 customers. More recently, the acquisition of Motivair in late 2024, a specialist in liquid cooling for AI infrastructure, directly positions Schneider at the heart of the hyperscaler buildout. The joint launch of AI-Ready Data Center Reference Designs with NVIDIA in October 2025 confirmed Schneider's ambition to become the infrastructure partner for the AI era. The target is clear: digital and recurring revenues to exceed 70% of group sales by 2030."),
        ("4. GOVERNANCE EVENT: THE CEO TRANSITION OF NOVEMBER 2024",
         "November 2024 brought an unexpected governance event: the Board removed CEO Peter Herweck after just 18 months, citing divergences in the execution of the company roadmap at a time of significant opportunities. His replacement, Olivier Blum, a 30-year Schneider veteran who had led the Energy Management division, was immediately seen as a reassuring choice. Under Blum's leadership, the 2025 Capital Markets Day set a compelling five-year ambition (Advancing Energy Tech), and 2025 results vindicated the change. Separately, the group was fined €207 million by French antitrust regulators in October 2024 for a price-fixing arrangement in the low-voltage electrical equipment market, a reputational headwind, but a manageable financial impact."),
        ("5. BALANCE SHEET &amp; LEVERAGE: A DELIBERATE BUT WATCHABLE INCREASE IN DEBT",
         "Schneider has grown significantly, and it has financed that growth. Net debt increased sharply in 2025, primarily to fund the full acquisition of its Indian subsidiary (SEIPL) from Temasek, a strategic bet on India as a core hub. The resulting Net Debt/EBITDA of 1.86x remains within investment-grade territory, but represents the highest leverage of the five-year period. Interest coverage, while still comfortable at 14x, has compressed significantly from the near-44x of 2021. The speed of this leverage build-up is the principal financial risk the Board must address in 2026: a clear deleveraging roadmap is now essential."),
        ("6. CASH FLOW: HIGH QUALITY AND SELF-FINANCING CAPACITY",
         "Schneider's cash generation is one of the strongest pillars of its investment case. Operating cash flow has grown by nearly +70% over five years, reaching €6.1 billion in 2025, with free cash flow conversion slightly above 100% of adjusted net income, confirming that earnings are real and cash-backed. The progressive dividend policy, maintained for 16 consecutive years, was reaffirmed with a dividend of €4.2 per share. Capex is rising in line with the group's manufacturing hub strategy, but remains fully self-financed. One area to watch: working capital needs are growing faster than revenue, and days receivables are creeping upward, a discipline issue at scale."),
        ("7. SHAREHOLDER VALUE: STRONG RETURNS, BUT VALUATION NOW STRETCHED",
         "Schneider's Total Shareholder Return grew +89% over the past three years, an exceptional performance. EPS and dividends have grown steadily throughout the period. The stock now trades at a premium valuation, reflecting the market's confidence in Schneider's positioning at the intersection of electrification, AI infrastructure and industrial automation. This premium, however, comes with expectations: the group must continue to deliver on margin expansion and organic growth targets. Any execution misstep, in AVEVA integration, Industrial Automation recovery, or leverage management, would likely trigger a swift market re-rating."),
    ]:
        st.markdown(f"""<div style="background:white;border:1px solid #E2E8F0;border-radius:8px;padding:14px 18px;margin-bottom:10px"><div style="font-size:10px;font-weight:700;color:{TEAL};text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">{_t}</div><p style="font-size:12px;color:#2D3748;line-height:1.7;margin:0">{_b}</p></div>""", unsafe_allow_html=True)

    section("Part II — Strategic Recommendations")
    for _pr, _rt, _rb in [
        ("CRITICAL", "Reallocate Capital Toward the Highest-ROCE Vertical",
         "The financial data speaks clearly: ROCE has improved from 8.3% to 12.1% over the period, but the group's capital allocation has not yet fully reflected where returns are highest. Data center infrastructure generates Schneider's best margins and fastest payback periods, yet capex is still spread across segments with structurally lower returns. The €388 million impairment booked in 2025 signals that not all deployed capital is working hard enough. Recommendation: formally reweight capex and M&amp;A spend toward high-ROCE verticals, and apply strict ROCE hurdle rates (minimum 12% by year 3) to all new investments. A disciplined capital rotation would accelerate group ROCE toward 15%+, directly supporting EV expansion and the current premium P/E multiple of 31x."),
        ("CRITICAL", "Ring-Fence R&amp;D Spend as a Fixed % of Revenue",
         "R&amp;D at 3.4% of revenue is structurally inconsistent with a 31x P/E valuation. The market is pricing Schneider as a technology platform, but the R&amp;D line looks like an industrial company. If this gap is not closed, the risk is a multiple de-rating that would destroy more shareholder value than any single operational miss. Committing to 5% R&amp;D/revenue (funded from the 1pp of SG&amp;A/sales savings already being generated) would directly protect AVEVA's ARR growth trajectory, strengthen gross margin over time as software mix increases, and reinforce the EV/EBITDA multiple. Every percentage point of gross margin expansion at €40bn of revenue is worth €400m in EBIT — the return on R&amp;D investment is not abstract, it is measurable."),
        ("HIGH", "Establish a Hard Leverage Ceiling and Deleverage by End-2027",
         "Net Debt/EBITDA rose from 1.07x to 1.86x in a single year, the sharpest leverage increase in the five-year period, while interest coverage compressed from 44x to 14x. In a higher-for-longer rate environment, each 50bps rise in refinancing costs on €15bn of debt represents €75m of additional annual interest, directly compressing net margin and EPS. The CFO must present a formal deleveraging roadmap to the Board: target Net Debt/EBITDA below 1.5x by end-2027, financed through organic FCF after dividends (~€2.4bn/year available). Any M&amp;A above €1bn must be subject to a Board-level leverage impact review with a hard 2.0x ceiling. Achieving this target would reduce net interest costs, expand net margin back toward 11–12%, and support a re-rating of the equity."),
        ("HIGH", "Protect the Dividend with a Formal Payout Policy",
         "The payout ratio reached 52.6% in 2025, the highest of the period, while reported net income declined slightly year-on-year. This creates a structural tension: the progressive dividend policy (16 consecutive years of growth) is a key pillar of the investment case, but it must not come at the expense of balance sheet flexibility. Recommendation: formalise a payout policy of 45–50% of adjusted net income, with a clear rule that excess FCF above that threshold is directed toward debt reduction when Net Debt/EBITDA exceeds 1.5x. This framework protects the dividend streak, reassures income investors, prevents EPS dilution from excessive leverage, and gives management a transparent capital allocation hierarchy to communicate to the market."),
        ("MEDIUM", "Unlock €400–500m of Cash Through Working Capital Discipline",
         "Working Capital Need has grown +127% over five years versus revenue growth of +39%, a clear execution gap. DSO has crept up to 88 days, DPO is healthy at 144 days but must be defended, and the CCC of 27 days, while excellent, is at risk of deteriorating as growth accelerates in higher-DSO markets (Middle East, India). Each day of DSO improvement at current revenue scale releases over €100m of cash. A 5-day DSO reduction target by end-2027 would free €550m, reducing net debt, cutting interest costs, and improving the Net Debt/EBITDA ratio by approximately 0.07x without any P&amp;L impact. Tools exist: supply chain financing, dynamic discounting, AR securitisation. The CFO should set quarterly DSO/DPO/CCC targets with Board-level reporting."),
        ("MEDIUM", "Revisit the M&amp;A Integration Framework to Protect Book Value",
         "Intangible assets at 50% of total assets is the single largest balance sheet risk. The €388m impairment in 2025, first of the period, is an early warning that acquired book values are not always preserved. With AVEVA (€11.9bn) and Motivair ($850m) both relatively recent, and Industrial Automation margins still at only 14.2% EBITDA, the integration execution risk is live. Recommendation: impose ROCE hurdle rates on every acquired entity (minimum 10% by year 3), run annual goodwill stress tests, and define automatic restructuring triggers if targets are missed. This protects book value per share, limits future impairment risk, and instils the financial discipline that ensures M&amp;A creates — rather than destroys — shareholder value."),
        ("MEDIUM", "Implement a Structural FX Hedging &amp; Revenue Mix Strategy",
         "FX is a silent but material drag on financial performance. In 2025, currency headwinds cost €410 million in cash flow impact, and management has already guided for a further €850–950 million revenue impact in FY2026, equivalent to roughly 2% of total revenues wiped out before a single product is sold. With the US dollar weakening against the euro and EM currencies volatile, this is not a transient issue. The CFO should implement a rolling 12-month FX hedging programme on the group's top 5 currency exposures (USD, CNY, INR, GBP, BRL), targeting 70–80% coverage of projected net exposures. In parallel, the regionalization strategy (90% local sourcing by end of cycle) creates a natural hedge — the Board should track natural hedge ratio as a formal KPI. Every 10% improvement in natural hedge coverage reduces P&amp;L FX sensitivity by an estimated €80–100m annually at current revenue scale."),
    ]:
        _pc = RED if _pr == "CRITICAL" else "#D97706" if _pr == "HIGH" else TEAL
        _pbg = "#FEF2F2" if _pr == "CRITICAL" else "#FFFBEB" if _pr == "HIGH" else "#F0FFF4"
        st.markdown(f"""<div style="background:white;border:1px solid #E2E8F0;border-left:4px solid {_pc};border-radius:0 8px 8px 0;padding:14px 18px;margin-bottom:10px"><div style="display:flex;align-items:center;gap:10px;margin-bottom:8px"><span style="background:{_pbg};color:{_pc};font-size:10px;font-weight:700;padding:3px 8px;border-radius:4px;letter-spacing:.05em">{_pr}</span><span style="font-size:11px;font-weight:700;color:{NAVY}">{_rt}</span></div><p style="font-size:12px;color:#2D3748;line-height:1.7;margin:0">{_rb}</p></div>""", unsafe_allow_html=True)

    section("Key Risks to Monitor")
    for _rk, _rl, _rr in [
        ("RISK 1", "Leverage &amp; Refinancing Risk",
         "Net Debt/EBITDA is approaching 1.9x, with interest coverage compressing rapidly. In a higher-for-longer rate environment, and with FX headwinds (management guided for a -€850–950 million revenue impact in FY2026), any earnings softness could push leverage above the 2.0x threshold. Bond maturities must be actively managed and the deleveraging roadmap enforced."),
        ("RISK 2", "Geopolitical &amp; Tariff Exposure",
         "Schneider operates globally and is exposed to US-China trade tensions, import tariffs and currency volatility. In 2025, the group was forced to accelerate pricing in North America to offset tariff costs, and expects flat-to-negative gross margin in H1 2026. With China remaining deflationary and Europe sluggish in several markets, geographic concentration of margin risk is a real consideration. The 90%-regional-sourcing target is the right structural answer, but will take years to fully implement."),
        ("RISK 3", "Cybersecurity &amp; Platform Vulnerability",
         "Schneider has been the target of repeated cybersecurity incidents. A ransomware attack in early 2024 directly impacted EcoStruxure Resource Advisor, used by over 2,000 companies globally. As the group's business model becomes more software-intensive and its platforms more deeply embedded in critical infrastructure, the cyber risk profile grows commensurately. A successful attack on EcoStruxure or AVEVA could cause severe operational disruption and reputational damage. Cybersecurity must receive board-level governance attention, not just IT-level management."),
        ("RISK 4", "Goodwill Impairment &amp; M&amp;A Integration",
         "With intangible assets representing 50% of total assets, Schneider carries a significant goodwill base from decades of acquisitions. The €388 million impairment recorded in 2025, a first in the period under review, is an early warning. The AVEVA integration, while progressing, must be closely tracked: any shortfall in ARR growth, margin expansion or cross-selling with EcoStruxure would risk a larger impairment charge. Annual goodwill stress tests should be a standard governance process."),
        ("RISK 5", "R&amp;D Underinvestment &amp; Software Disruption",
         "The structural gap in R&amp;D versus technology peers is not just a strategic concern, it is a valuation risk. If Schneider is perceived as primarily a hardware company with a software overlay, it will trade at industrial multiples rather than technology multiples. The pace of AI development makes this risk more acute: new entrants could commoditize elements of AVEVA's platform faster than anticipated. The Board must treat R&amp;D investment as a non-negotiable line item, not a variable cost."),
    ]:
        st.markdown(f"""<div style="background:white;border:1px solid #E2E8F0;border-left:4px solid {RED};border-radius:0 8px 8px 0;padding:14px 18px;margin-bottom:10px"><div style="display:flex;align-items:center;gap:10px;margin-bottom:8px"><span style="background:#FEF2F2;color:{RED};font-size:10px;font-weight:700;padding:3px 8px;border-radius:4px;letter-spacing:.05em">{_rk}</span><span style="font-size:11px;font-weight:700;color:{NAVY}">{_rl}</span></div><p style="font-size:12px;color:#2D3748;line-height:1.7;margin:0">{_rr}</p></div>""", unsafe_allow_html=True)

# ─── TAB 7: DATA ENTRY ─────────────────────────────────────────────
with tabs[6]:
    section("Manual Data Entry")
    st.markdown(
        '<p style="font-size:13px;color:#4B5563;margin-bottom:16px">'
        'Enter raw financial figures for a new fiscal year. '
        'All charts and trend analysis update automatically once you submit.</p>',
        unsafe_allow_html=True
    )

    col_yr, col_clr = st.columns([3, 1])
    with col_yr:
        entry_year = st.selectbox("Fiscal Year to Add / Update",
                                  [2026, 2027, 2028, 2029, 2030])
    with col_clr:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("Clear All Entries", type="secondary"):
            st.session_state.manual_entries = {}
            st.rerun()

    with st.form("data_entry_form"):

        section("Income Statement (EURm)")
        c1, c2, c3 = st.columns(3)
        with c1:
            inp_revenue = st.number_input("Revenue",      value=0.0, step=100.0, format="%.1f")
            inp_gross   = st.number_input("Gross Profit", value=0.0, step=100.0, format="%.1f")
            inp_ebita   = st.number_input("EBITA",        value=0.0, step=100.0, format="%.1f")
        with c2:
            inp_ebitda  = st.number_input("EBITDA",       value=0.0, step=100.0, format="%.1f")
            inp_ebit    = st.number_input("EBIT",         value=0.0, step=100.0, format="%.1f")
            inp_ni      = st.number_input("Net Income",   value=0.0, step=100.0, format="%.1f")
        with c3:
            inp_cogs    = st.number_input("COGS (enter as negative)",             value=0.0, step=100.0, format="%.1f")
            inp_int_exp = st.number_input("Gross Interest Expense (as negative)", value=0.0, step=10.0,  format="%.1f")

        section("Cash Flow Statement (EURm)")
        c1, c2, c3 = st.columns(3)
        with c1:
            inp_cfo     = st.number_input("Cash from Operations (CFO)",  value=0.0, step=100.0, format="%.1f")
            inp_capex   = st.number_input("Net CapEx (as negative)",     value=0.0, step=50.0,  format="%.1f")
            inp_cfi     = st.number_input("Cash from Investing (CFI)",   value=0.0, step=100.0, format="%.1f")
        with c2:
            inp_cff     = st.number_input("Cash from Financing (CFF)",   value=0.0, step=100.0, format="%.1f")
            inp_divs    = st.number_input("Dividends (as negative)",     value=0.0, step=50.0,  format="%.1f")
        with c3:
            inp_buyback = st.number_input("Buybacks (as negative)",      value=0.0, step=50.0,  format="%.1f")
            inp_ma      = st.number_input("M&A / Acquisitions",          value=0.0, step=100.0, format="%.1f")

        section("Balance Sheet Ratios")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            inp_roe      = st.number_input("ROE (%)",               value=0.0, step=0.1,  format="%.2f")
            inp_roce     = st.number_input("ROCE (%)",              value=0.0, step=0.1,  format="%.2f")
            inp_asset_to = st.number_input("Asset Turnover (x)",   value=0.0, step=0.01, format="%.3f")
            inp_eq_mult  = st.number_input("Equity Multiplier (x)",value=0.0, step=0.1,  format="%.2f")
        with c2:
            inp_cur  = st.number_input("Current Ratio (x)", value=0.0, step=0.01, format="%.2f")
            inp_qck  = st.number_input("Quick Ratio (x)",   value=0.0, step=0.01, format="%.2f")
            inp_csh  = st.number_input("Cash Ratio (x)",    value=0.0, step=0.01, format="%.2f")
        with c3:
            inp_net_debt = st.number_input("Net Debt (EURm)",        value=0.0, step=100.0, format="%.1f")
            inp_de       = st.number_input("Debt / Equity (x)",      value=0.0, step=0.1,   format="%.2f")
            inp_debt_r   = st.number_input("Total Debt Ratio (%)",   value=0.0, step=0.1,   format="%.2f")
            inp_int_cov  = st.number_input("Int. Coverage EBIT (x)", value=0.0, step=0.1,   format="%.1f")
        with c4:
            inp_dso = st.number_input("DSO (days)", value=0.0, step=1.0, format="%.1f")
            inp_dio = st.number_input("DIO (days)", value=0.0, step=1.0, format="%.1f")
            inp_dpo = st.number_input("DPO (days)", value=0.0, step=1.0, format="%.1f")

        submitted = st.form_submit_button("Add to Dashboard", type="primary")

        if submitted:
            if inp_revenue <= 0:
                st.error("Revenue must be greater than 0 to add an entry.")
            else:
                _fcf = inp_cfo + inp_capex
                _dda = inp_ebitda - inp_ebita
                _entry = {
                    "revenue":     inp_revenue,
                    "cogs":        inp_cogs,
                    "gross_profit":inp_gross,
                    "ebita":       inp_ebita,
                    "ebitda":      inp_ebitda,
                    "ebit":        inp_ebit,
                    "net_income":  inp_ni,
                    "int_expense": inp_int_exp,
                    "roe":         inp_roe  / 100,
                    "ros":         inp_ni   / inp_revenue if inp_revenue else None,
                    "asset_to":    inp_asset_to,
                    "eq_mult":     inp_eq_mult,
                    "gross_m":     inp_gross   / inp_revenue if inp_revenue else None,
                    "ebita_m":     inp_ebita   / inp_revenue if inp_revenue else None,
                    "ebit_m":      inp_ebit    / inp_revenue if inp_revenue else None,
                    "ebitda_m":    inp_ebitda  / inp_revenue if inp_revenue else None,
                    "dso":         inp_dso,
                    "dio":         inp_dio,
                    "dpo":         inp_dpo,
                    "ccc":         inp_dso + inp_dio - inp_dpo,
                    "current_r":   inp_cur,
                    "quick_r":     inp_qck,
                    "cash_r":      inp_csh,
                    "de_ratio":    inp_de,
                    "debt_ratio":  inp_debt_r / 100,
                    "int_cov":     inp_int_cov,
                    "nd_ebita":    inp_net_debt / inp_ebita  if inp_ebita  else None,
                    "payout":      None, "div_yield": None, "eps": None,
                    "pe":          None, "pb":         None,
                    "net_debt":    inp_net_debt,
                    "ev":          None, "ev_ebita":   None,
                    "roce":        inp_roce / 100,
                    "nd_ebitda":   inp_net_debt / inp_ebitda  if inp_ebitda  else None,
                    "int_cov_ebd": inp_ebitda   / abs(inp_int_exp) if inp_int_exp else None,
                    "cfo":         inp_cfo,
                    "capex":       inp_capex,
                    "cfi":         inp_cfi,
                    "cff":         inp_cff,
                    "divs":        inp_divs,
                    "buyback":     inp_buyback,
                    "ma":          inp_ma,
                    "fcf":         _fcf,
                    "fcf_m":       _fcf  / inp_revenue if inp_revenue else None,
                    "fcf_ni":      _fcf  / inp_ni      if inp_ni      else None,
                    "cfo_ebd":     inp_cfo / inp_ebitda if inp_ebitda else None,
                    "capex_r":     abs(inp_capex) / inp_revenue if inp_revenue else None,
                    "capex_da":    abs(inp_capex) / _dda        if _dda       else None,
                }
                st.session_state.manual_entries[entry_year] = _entry
                st.rerun()

    # ── Active entries summary ────────────────────────────────────────
    if st.session_state.manual_entries:
        section("Active Manual Entries")
        for _yr, _e in sorted(st.session_state.manual_entries.items()):
            st.markdown(
                f'<div style="background:white;border:1px solid #E2E8F0;border-radius:10px;'
                f'padding:14px 18px;margin-bottom:10px">'
                f'<span style="font-size:13px;font-weight:700;color:{NAVY}">FY{_yr}</span>'
                f'<span style="margin-left:16px;font-size:12px;color:{DGRAY}">'
                f'Revenue: <b>{safe(_e.get("revenue"),"eur")} €m</b> · '
                f'EBIT Margin: <b>{safe(_e.get("ebit_m"),"pct")}</b> · '
                f'FCF: <b>{safe(_e.get("fcf"),"eur")} €m</b> · '
                f'Net Debt/EBITDA: <b>{safe(_e.get("nd_ebitda"),"x")}</b> · '
                f'ROE: <b>{safe(_e.get("roe"),"pct")}</b>'
                f'</span></div>',
                unsafe_allow_html=True
            )
            if st.button(f"Remove FY{_yr}", key=f"del_{_yr}"):
                del st.session_state.manual_entries[_yr]
                st.rerun()
