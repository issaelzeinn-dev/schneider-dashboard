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
    path = "Schneider_Dashboard.xlsm"

    r  = pd.read_excel(path, sheet_name="Ratios",              header=None, engine="openpyxl")
    i  = pd.read_excel(path, sheet_name="Income Statement",    header=None, engine="openpyxl")
    cf = pd.read_excel(path, sheet_name="Cash Flow Statement", header=None, engine="openpyxl")
    bs = pd.read_excel(path, sheet_name="Balance Sheet",       header=None, engine="openpyxl")

    # IS cols ordered FY2021→FY2025: 17,14,11,8,5
    IC = [17, 14, 11, 8, 5]
    # Ratios cols ordered FY2021→FY2025: 7,6,5,4,3
    RC = [7, 6, 5, 4, 3]
    # CFS cols ordered FY2021→FY2025: 12,10,8,6,4
    CC = [12, 10, 8, 6, 4]

    def irow(df, c): return [df.iloc[c, x] for x in IC]
    def rrow(df, c): return [df.iloc[c, x] for x in RC]
    def crow(df, c): return [df.iloc[c, x] for x in CC]

    d = {}

    # Income Statement (row index = excel row - 1, 0-based, header is row 0)
    d["revenue"]      = irow(i, 2)   # 40152 25
    d["cogs"]         = irow(i, 3)
    d["gross_profit"] = irow(i, 4)
    d["ebita"]        = irow(i, 10)
    d["ebitda"]       = irow(i, 11)
    d["ebit"]         = irow(i, 13)
    d["net_income"]   = irow(i, 23)
    d["int_expense"]  = irow(i, 15)  # gross cost fin debt

    # Ratios (0-based)
    d["roe"]          = rrow(r, 3)
    d["ros"]          = rrow(r, 4)
    d["asset_to"]     = rrow(r, 5)
    d["eq_mult"]      = rrow(r, 6)
    d["gross_m"]      = rrow(r, 10)
    d["ebita_m"]      = rrow(r, 12)
    d["ebit_m"]       = rrow(r, 13)
    d["dso"]          = rrow(r, 15)
    d["dio"]          = rrow(r, 16)
    d["dpo"]          = rrow(r, 17)
    d["ccc"]          = rrow(r, 18)
    d["current_r"]    = rrow(r, 22)
    d["quick_r"]      = rrow(r, 23)
    d["cash_r"]       = rrow(r, 24)
    d["de_ratio"]     = rrow(r, 26)
    d["debt_ratio"]   = rrow(r, 27)
    d["int_cov"]      = rrow(r, 28)
    d["nd_ebita"]     = rrow(r, 29)
    d["payout"]       = rrow(r, 31)
    d["div_yield"]    = rrow(r, 32)
    d["eps"]          = rrow(r, 33)
    d["pe"]           = rrow(r, 34)
    d["pb"]           = rrow(r, 35)
    d["net_debt"]     = rrow(r, 36)
    d["ev"]           = rrow(r, 37)
    d["ev_ebita"]     = rrow(r, 38)
    d["roce"]         = rrow(r, 39)

    # Derived EBITDA ratios
    d["ebitda_m"]     = [e/rev if rev else None for e,rev in zip(d["ebitda"], d["revenue"])]
    d["nd_ebitda"]    = [nd/eb if eb else None for nd,eb in zip(d["net_debt"], d["ebitda"])]
    d["int_cov_ebd"]  = [eb/abs(ie) if ie else None for eb,ie in zip(d["ebitda"], d["int_expense"])]

    # Cash Flow Statement (0-based row)
    d["cfo"]     = crow(cf, 19)   # TOTAL I
    d["capex"]   = crow(cf, 23)   # Net capital expenditure (negative)
    d["cfi"]     = crow(cf, 28)   # TOTAL II
    d["cff"]     = crow(cf, 38)   # TOTAL III
    d["divs"]    = crow(cf, 36)   # Dividends SE shareholders
    d["buyback"] = crow(cf, 31)   # Treasury shares
    d["ma"]      = crow(cf, 24)   # Acquisitions net

    d["fcf"]     = [c + cap for c,cap in zip(d["cfo"], d["capex"])]
    d["fcf_m"]   = [f/rev if rev else None for f,rev in zip(d["fcf"], d["revenue"])]
    d["fcf_ni"]  = [f/ni if ni else None for f,ni in zip(d["fcf"], d["net_income"])]
    d["cfo_ebd"] = [c/e if e else None for c,e in zip(d["cfo"], d["ebitda"])]
    d["capex_r"] = [abs(cap)/rev if rev else None for cap,rev in zip(d["capex"], d["revenue"])]
    dda = [e - eb for e,eb in zip(d["ebitda"], d["ebita"])]
    d["capex_da"]= [abs(cap)/da if da else None for cap,da in zip(d["capex"], dda)]

    # Balance Sheet
    d["intangibles"]  = [bs.iloc[2,  3], bs.iloc[2,  6]]
    d["ppe"]          = [bs.iloc[3,  3], bs.iloc[3,  6]]
    d["oth_nca"]      = [bs.iloc[4,  3], bs.iloc[4,  6]]
    d["inventories"]  = [bs.iloc[6,  3], bs.iloc[6,  6]]
    d["receivables"]  = [bs.iloc[7,  3], bs.iloc[7,  6]]
    d["cash"]         = [bs.iloc[9,  3], bs.iloc[9,  6]]
    d["total_assets"] = [bs.iloc[11, 3], bs.iloc[11, 6]]
    d["equity"]       = [bs.iloc[16, 3], bs.iloc[16, 6]]
    d["lt_debt"]      = [bs.iloc[17, 3], bs.iloc[17, 6]]
    d["st_debt"]      = [bs.iloc[22, 3], bs.iloc[22, 6]]

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
        textinfo="percent", textfont=dict(size=10, color="#111111"),
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
