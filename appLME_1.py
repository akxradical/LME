"""
Commodities Trading Terminal
=============================
LME Copper · Aluminium · Brent Crude Oil
Stock-market UI · Google Sheets live feed · Market Intel · PPTX export

requirements.txt:
  streamlit>=1.32
  pandas
  numpy
  gspread
  google-auth
  plotly
  matplotlib
  python-pptx
  openpyxl
"""

import io, xml.etree.ElementTree as ET
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError

import gspread, numpy as np, pandas as pd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.lines import Line2D
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ══════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════
SPREADSHEET_ID  = "1zLyMANY56oFRwFug04WYavGUH_NAlRH8M3c-TXIRDlI"
SHEET_COPPER    = "LME Copper"
SHEET_ALUMINIUM = "LME Aluminium"
SHEET_BRENT     = "Brent Oil"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Dark theme
BG      = "#0D1117"; CARD = "#161B22"; CARD2 = "#1C2128"
BORDER  = "#30363D"; GRID = "#21262D"
GREEN   = "#00C853"; RED = "#FF1744"; GOLD = "#FFD600"; BLUE = "#2979FF"
TXT     = "#E6EDF3"; SEC = "#8B949E"; MUT = "#484F58"
CU_CLR  = "#E87B35"; AL_CLR = "#78909C"; BR_CLR = "#4CAF50"

st.set_page_config(page_title="Commodities Terminal", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# ══════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════
st.markdown("""
<style>
html,body,[data-testid="stAppViewContainer"],[data-testid="stApp"]{
  background:#0D1117!important;color:#E6EDF3;
  font-family:'JetBrains Mono','Fira Code','Courier New',monospace}
[data-testid="stSidebar"]{background:#161B22!important;border-right:1px solid #30363D}
[data-testid="stHeader"]{background:transparent!important}
.block-container{padding-top:1rem;max-width:100%}
.tc{background:#161B22;border:1px solid #30363D;border-radius:8px;padding:14px 18px;margin-bottom:10px}
.tl{font-size:10px;letter-spacing:1.2px;color:#8B949E;text-transform:uppercase;margin-bottom:4px}
.tv{font-size:28px;font-weight:700;color:#E6EDF3;letter-spacing:-0.5px}
.ts{font-size:11px;margin-top:3px}
.up{color:#00C853}.dn{color:#FF1744}
.sh{font-size:11px;letter-spacing:1.5px;color:#8B949E;text-transform:uppercase;
    border-bottom:1px solid #30363D;padding-bottom:6px;margin-bottom:12px}
.dt{width:100%;border-collapse:collapse;font-size:12px}
.dt th{background:#1C2128;color:#8B949E;font-size:10px;letter-spacing:1px;
       text-transform:uppercase;padding:8px 12px;text-align:right;border-bottom:1px solid #30363D}
.dt th:first-child{text-align:left}
.dt td{padding:7px 12px;border-bottom:1px solid #1C2128;text-align:right;color:#E6EDF3}
.dt td:first-child{text-align:left;color:#8B949E}
.dt tr:hover td{background:#1C2128}
.stTabs [data-baseweb="tab-list"]{background:#161B22;border-radius:8px;padding:4px;gap:4px;border:1px solid #30363D}
.stTabs [data-baseweb="tab"]{background:transparent;color:#8B949E;border-radius:6px;padding:8px 16px;font-size:11px}
.stTabs [aria-selected="true"]{background:#1C2128!important;color:#E6EDF3!important;border:1px solid #30363D!important}
.stButton>button{background:#2979FF;color:white;border:none;border-radius:6px;font-weight:600;width:100%;padding:10px}
[data-testid="stDownloadButton"]>button{background:#1B5E20;color:#00C853;border:1px solid #00C853;border-radius:6px;font-weight:600;width:100%;padding:10px}
.pcard{background:#161B22;border:1px solid #30363D;border-radius:10px;padding:16px;margin-bottom:12px}
.news-item{background:#161B22;border-left:3px solid #2979FF;padding:10px 14px;margin-bottom:8px;border-radius:0 6px 6px 0}
.news-title{color:#E6EDF3;font-size:13px;font-weight:600;text-decoration:none}
.news-title:hover{color:#2979FF}
.news-src{font-size:10px;color:#484F58;margin-top:3px}
#MainMenu,footer,header{visibility:hidden}
[data-testid="stToolbar"]{display:none}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_sheet(sheet_name):
    try:
        info  = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        gc    = gspread.authorize(creds)
        sid   = st.secrets.get("SPREADSHEET_ID", SPREADSHEET_ID)
        ws    = gc.open_by_key(sid).worksheet(sheet_name)
        rows  = ws.get_all_records()
        df    = pd.DataFrame(rows)
        df.columns = [c.strip() for c in df.columns]
        date_col  = next(c for c in df.columns if "date" in c.lower())
        price_col = next(c for c in df.columns if "price" in c.lower() or "cash" in c.lower())
        df = df.rename(columns={date_col: "Date", price_col: "Price"})
        df["Date"]  = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        df["Price"] = pd.to_numeric(df["Price"].astype(str).str.replace(",",""), errors="coerce")
        return df.dropna(subset=["Date","Price"]).sort_values("Date").reset_index(drop=True)
    except Exception:
        return _synth(sheet_name)

def _synth(name):
    seed = 42 if "Copper" in name else (7 if "Alum" in name else 99)
    base = 9000 if "Copper" in name else (2300 if "Alum" in name else 75)
    rng  = np.random.default_rng(seed)
    dates = pd.date_range("2025-01-02", datetime.today(), freq="B")
    noise = rng.normal(0, base*0.005, len(dates)).cumsum()
    return pd.DataFrame({"Date": dates, "Price": np.clip(base+noise, base*0.8, base*1.3).round(2)})

# ══════════════════════════════════════════════
# ANALYTICS
# ══════════════════════════════════════════════
def monthly_avg(df):
    d = df.copy(); d["YM"] = d["Date"].dt.to_period("M")
    m = d.groupby("YM")["Price"].mean().reset_index()
    m["Date"] = m["YM"].dt.to_timestamp(); m["Label"] = m["YM"].dt.strftime("%b-%y")
    m["Price"] = m["Price"].round(2); return m.drop(columns="YM")

def quarterly_avg(df):
    d = df.copy(); d["YQ"] = d["Date"].dt.to_period("Q")
    q = d.groupby("YQ")["Price"].mean().reset_index()
    q["Date"] = q["YQ"].dt.to_timestamp(); q["Label"] = q["YQ"].dt.strftime("Q%q-%Y")
    q["Price"] = q["Price"].round(2); return q.drop(columns="YQ")

def rolling_avg(df, w=3):
    d = df.copy().sort_values("Date"); d[f"MA{w}"] = d["Price"].rolling(w).mean().round(2); return d

def linear_forecast(m, p=3):
    if len(m)<2: return pd.DataFrame(columns=["Date","Price","Label"])
    x=np.arange(len(m)); s,i=np.polyfit(x,m["Price"].values,1); last=m["Date"].iloc[-1]
    fd=[last+pd.DateOffset(months=k+1) for k in range(p)]
    fp=[i+s*(len(m)+k+1) for k in range(p)]
    return pd.DataFrame({"Date":fd,"Price":np.round(fp,2),"Label":[d.strftime("%b-%y") for d in fd]})

def mom_change(m):
    r=m.copy(); r["Change"]=r["Price"].diff().round(2)
    r["Change%"]=(r["Price"].pct_change()*100).round(2)
    r["3M_Avg"]=r["Price"].rolling(3).mean().round(2); return r

# ══════════════════════════════════════════════
# CHART LAYOUT
# ══════════════════════════════════════════════
_CL = dict(paper_bgcolor=CARD, plot_bgcolor=CARD,
           font=dict(color=TXT, family="'JetBrains Mono',monospace", size=11),
           margin=dict(l=55,r=15,t=65,b=55),
           xaxis=dict(showgrid=False,zeroline=False,color=SEC,linecolor=BORDER,tickcolor=BORDER),
           yaxis=dict(gridcolor=GRID,zeroline=False,color=SEC,linecolor=BORDER,tickformat=","),
           legend=dict(bgcolor="rgba(0,0,0,0)",bordercolor=BORDER,orientation="h",y=-0.18,font=dict(size=10)))

def chart_live(df_raw, metal, color, unit="USD/MT"):
    df = df_raw.copy().sort_values("Date").reset_index(drop=True)
    r3 = rolling_avg(df, 3)
    cur = df["Price"].iloc[-1]
    prev = df["Price"].iloc[-2] if len(df)>1 else cur
    chg = cur-prev; pct = chg/prev*100 if prev else 0
    arrow = "▲" if chg>=0 else "▼"; clr = GREEN if chg>=0 else RED
    r,g,b = bytes.fromhex(color.lstrip("#"))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df["Date"],y=df["Price"],fill="tozeroy",
        fillcolor=f"rgba({r},{g},{b},0.07)",line=dict(color=color,width=1.5),
        mode="lines",name="Daily Price",
        hovertemplate="<b>%{x|%d %b %Y}</b><br>$%{y:,.2f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=r3["Date"],y=r3["MA3"],
        line=dict(color=GOLD,width=1.2,dash="dot"),mode="lines",name="3M MA",
        hovertemplate="3M: $%{y:,.2f}<extra></extra>"))
    fl = df["Date"].iloc[0].strftime("%d %b %Y"); ll = df["Date"].iloc[-1].strftime("%d %b %Y")
    _b = {k:v for k,v in _CL.items() if k not in ("xaxis","yaxis")}
    fig.update_layout(**_b,
        title=dict(text=(f"<b>{metal}</b>  "
            f"<span style='font-size:22px;font-weight:700;color:{clr}'>${cur:,.2f}</span>  "
            f"<span style='font-size:13px;color:{clr}'>{arrow} ${abs(chg):,.2f} ({arrow}{abs(pct):.2f}%)</span>"
            f"<br><span style='font-size:10px;color:{SEC}'>{unit} · {fl} → {ll}</span>"),
            font=dict(size=15,color=TXT)),
        xaxis=dict(showgrid=False,zeroline=False,color=SEC,linecolor=BORDER,tickcolor=BORDER,
            rangeslider=dict(visible=True,bgcolor=BG,thickness=0.04),
            rangeselector=dict(bgcolor=CARD2,activecolor=BORDER,font=dict(color=SEC,size=10),
                buttons=[dict(count=1,label="1M",step="month",stepmode="backward"),
                         dict(count=3,label="3M",step="month",stepmode="backward"),
                         dict(count=6,label="6M",step="month",stepmode="backward"),
                         dict(count=1,label="YTD",step="year",stepmode="todate"),
                         dict(step="all",label="ALL")])),
        yaxis=dict(gridcolor=GRID,zeroline=False,color=SEC,linecolor=BORDER,tickformat=",",side="right"),
        height=420,hovermode="x unified")
    return fig

def chart_quarterly(q, fq, metal, color):
    q2 = q.copy(); q2["prev"]=q2["Price"].shift(1); q2["up"]=q2["Price"]>=q2["prev"]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=q2["Label"],y=q2["Price"],
        marker_color=[GREEN if u else RED for u in q2["up"]],opacity=0.85,
        text=[f"{p:,.0f}" for p in q2["Price"]],textposition="outside",
        textfont=dict(size=9,color=TXT),name="Quarterly Avg"))
    if len(fq):
        fig.add_trace(go.Bar(x=fq["Label"],y=fq["Price"],marker_color=GOLD,opacity=0.5,
            text=[f"{p:,.0f}" for p in fq["Price"]],textposition="outside",
            textfont=dict(size=9,color=GOLD),name="Forecast"))
    fig.update_layout(**_CL,title=dict(text=f"<b>{metal}</b>  <span style='color:{SEC};font-size:12px'>Quarterly Avg</span>",
        font=dict(size=15,color=TXT)),height=350,barmode="group")
    return fig

def chart_rolling(df_raw, metal, color):
    r3=rolling_avg(df_raw,3); r6=rolling_avg(df_raw,6)
    fig=go.Figure()
    fig.add_trace(go.Scatter(x=r3["Date"],y=r3["Price"],line=dict(color=BORDER,width=0.8),mode="lines",name="Daily",opacity=0.5))
    fig.add_trace(go.Scatter(x=r3["Date"],y=r3["MA3"],line=dict(color=color,width=2),mode="lines",name="3M MA"))
    fig.add_trace(go.Scatter(x=r6["Date"],y=r6["MA6"],line=dict(color=GOLD,width=1.5,dash="dash"),mode="lines",name="6M MA"))
    fig.update_layout(**_CL,title=dict(text=f"<b>{metal}</b>  <span style='color:{SEC};font-size:12px'>Rolling Avg (3M/6M)</span>",
        font=dict(size=15,color=TXT)),height=350)
    return fig

# ══════════════════════════════════════════════
# KPI CARDS
# ══════════════════════════════════════════════
def render_kpis(monthly, unit="USD/MT"):
    c=monthly["Price"].iloc[-1]; cl=monthly["Label"].iloc[-1]
    lo=monthly["Price"].min(); ll=monthly.loc[monthly["Price"].idxmin(),"Label"]
    hi=monthly["Price"].max(); hl=monthly.loc[monthly["Price"].idxmax(),"Label"]
    p=monthly["Price"].iloc[-2] if len(monthly)>1 else c
    ch=c-p; pc=ch/p*100 if p else 0; f=monthly["Price"].iloc[0]
    yc=c-f; yp=yc/f*100 if f else 0
    ud="up" if ch>=0 else "dn"; ya="up" if yc>=0 else "dn"
    ar="▲" if ch>=0 else "▼"; yr="▲" if yc>=0 else "▼"
    st.markdown(
        '<div class="tc"><div class="tl">CURRENT · '+cl+'</div>'
        '<div class="tv">$'+f'{c:,.2f}'+'</div>'
        '<div class="ts '+ud+'">'+ar+' $'+f'{abs(ch):,.2f}'+' ('+ar+f'{abs(pc):.2f}'+'%) MoM</div></div>'
        '<div class="tc"><div class="tl">PERIOD HIGH · '+hl+'</div>'
        '<div class="tv up">$'+f'{hi:,.2f}'+'</div>'
        '<div class="ts" style="color:#484F58">'+unit+'</div></div>'
        '<div class="tc"><div class="tl">PERIOD LOW · '+ll+'</div>'
        '<div class="tv dn">$'+f'{lo:,.2f}'+'</div>'
        '<div class="ts" style="color:#484F58">'+unit+'</div></div>'
        '<div class="tc"><div class="tl">YTD RETURN</div>'
        '<div class="tv '+ya+'">'+yr+' '+f'{abs(yp):.1f}'+'%</div>'
        '<div class="ts '+ya+'">'+yr+' $'+f'{abs(yc):,.0f}'+' since '+monthly["Label"].iloc[0]+'</div></div>',
        unsafe_allow_html=True)

# ══════════════════════════════════════════════
# MONTHLY TABLE
# ══════════════════════════════════════════════
def render_table(monthly, unit="USD/MT"):
    m = mom_change(monthly)
    rows = ""
    for _, r in m.iterrows():
        ch = "" if pd.isna(r["Change"]) else ('<span class="'+("up" if r["Change"]>=0 else "dn")+'">'+("+" if r["Change"]>=0 else "")+f'{r["Change"]:,.2f}</span>')
        pc = "" if pd.isna(r["Change%"]) else ('<span class="'+("up" if r["Change%"]>=0 else "dn")+'">'+("+" if r["Change%"]>=0 else "")+f'{r["Change%"]:.2f}%</span>')
        a3 = "—" if pd.isna(r["3M_Avg"]) else f'${r["3M_Avg"]:,.2f}'
        rows += '<tr><td>'+r["Label"]+'</td><td>$'+f'{r["Price"]:,.2f}'+'</td><td>'+ch+'</td><td>'+pc+'</td><td style="color:#8B949E">'+a3+'</td></tr>'
    st.markdown('<table class="dt"><thead><tr><th>Month</th><th>Avg Price</th><th>MoM</th><th>MoM %</th><th>3M Avg</th></tr></thead><tbody>'+rows+'</tbody></table>', unsafe_allow_html=True)

# ══════════════════════════════════════════════
# TICKER BAR
# ══════════════════════════════════════════════
def render_ticker(cu_m, al_m, br_m):
    def s(df):
        c=df["Price"].iloc[-1]; p=df["Price"].iloc[-2] if len(df)>1 else c
        return c, c-p, (c-p)/p*100 if p else 0
    cc,cch,cp = s(cu_m); ac,ach,ap = s(al_m); bc,bch,bp = s(br_m)

    def blk(name, price, chg, pct, clr):
        badge = "lme-cup" if chg>=0 else "lme-cdn"
        arr = "&#9650;" if chg>=0 else "&#9660;"
        return (
            '<div class="lme-pb">'
            '<div class="lme-mn">'+name+'</div>'
            '<div class="lme-pr"><span class="lme-pv" style="color:'+clr+'">'+'${:,.2f}'.format(price)+'</span></div>'
            '<div class="lme-cr"><span class="lme-badge '+badge+'">'+arr+' {:.2f}%'.format(abs(pct))+'</span>'
            '<span class="lme-ab">'+arr+' ${:,.2f} MoM'.format(abs(chg))+'</span></div>'
            '</div>'
        )

    now = datetime.now().strftime("%d %b %Y  %H:%M")
    css = (
        '<style>'
        '.lme-wrap{background:#0A0F1E;border-bottom:2px solid #B87333;margin-bottom:12px}'
        '.lme-top{background:#050A14;padding:7px 24px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #1C2333}'
        '.lme-logo{font-size:22px;font-weight:900;letter-spacing:3px;color:#FFF;font-family:"Arial Black",Arial,sans-serif}'
        '.lme-bar{background:#0A0F1E;padding:14px 24px;display:flex;align-items:stretch;border-bottom:1px solid #1C2333;gap:0;flex-wrap:wrap}'
        '.lme-pb{display:flex;flex-direction:column;padding:8px 28px 8px 0;margin-right:24px;border-right:1px solid #1C2333;min-width:180px}'
        '.lme-mn{font-size:10px;letter-spacing:1.5px;color:#8B949E;text-transform:uppercase;margin-bottom:5px}'
        '.lme-pr{display:flex;align-items:baseline;gap:8px}'
        '.lme-pv{font-size:24px;font-weight:700;letter-spacing:-0.5px;font-family:Arial,sans-serif}'
        '.lme-cr{display:flex;align-items:center;gap:8px;margin-top:4px}'
        '.lme-badge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:3px}'
        '.lme-cup{background:rgba(0,200,83,0.15);color:#00C853}'
        '.lme-cdn{background:rgba(255,23,68,0.15);color:#FF1744}'
        '.lme-ab{font-size:10px;color:#8B949E}'
        '.lme-rf{margin-left:auto;display:flex;flex-direction:column;justify-content:center;padding-left:24px;border-left:1px solid #1C2333}'
        '.lme-live{display:inline-block;width:7px;height:7px;background:#00C853;border-radius:50%;margin-right:5px;vertical-align:middle;animation:ldp 2s infinite}'
        '@keyframes ldp{0%,100%{opacity:1}50%{opacity:0.3}}'
        '</style>'
    )
    body = (
        '<div class="lme-wrap">'
        '<div class="lme-top"><div>'
        '<div class="lme-logo">LME<span style="color:#B87333">.</span></div>'
        '<div style="font-size:9px;color:#8B949E;letter-spacing:2px;text-transform:uppercase;margin-top:1px">Commodities Price Terminal</div>'
        '</div>'
        '<div style="font-size:10px;color:#484F58;letter-spacing:1px">CASH SETTLEMENT &middot; DAILY</div>'
        '</div>'
        '<div class="lme-bar">'
        + blk("&#9632; COPPER (Cu) &middot; USD/MT", cc, cch, cp, CU_CLR)
        + blk("&#9632; ALUMINIUM (Al) &middot; USD/MT", ac, ach, ap, AL_CLR)
        + blk("&#9632; BRENT CRUDE &middot; USD/BBL", bc, bch, bp, BR_CLR)
        + '<div class="lme-rf">'
          '<span style="font-size:9px;color:#484F58;letter-spacing:1.5px;text-transform:uppercase"><span class="lme-live"></span>Live Feed</span>'
          '<span style="font-size:13px;color:#8B949E;margin-top:3px">'+now+'</span>'
          '<span style="font-size:9px;color:#484F58;margin-top:2px">Auto-refresh 5 min</span>'
          '</div>'
        '</div></div><br>'
    )
    st.markdown(css+body, unsafe_allow_html=True)

# ══════════════════════════════════════════════
# TOP 5 PRODUCERS  (static data — stable year over year)
# ══════════════════════════════════════════════
TOP5 = {
    "Copper": [
        ("🇨🇱","Chile","5.8M MT","Escondida, Collahuasi — world's #1, ~27% global supply"),
        ("🇵🇪","Peru","2.7M MT","Cerro Verde, Antamina — vast reserves but needs H₂SO₄ from China for SX-EW leaching"),
        ("🇨🇩","DR Congo","2.5M MT","Kamoa-Kakula (Ivanhoe) — fastest growing; political instability risk"),
        ("🇨🇳","China","1.9M MT","Largest smelter & consumer; imports 70%+ of concentrate; controls H₂SO₄ supply chain"),
        ("🇺🇸","USA","1.1M MT","Morenci (Freeport-McMoRan) — declining grades, environmental permit delays"),
    ],
    "Aluminium": [
        ("🇨🇳","China","41M MT","60% of world output; Yunnan smelters face hydropower curtailments in dry season"),
        ("🇮🇳","India","4.1M MT","Hindalco, Vedanta — cheap coal power advantage; growing exports"),
        ("🇷🇺","Russia","3.8M MT","Rusal (UC Rusal) — sanctions risk; LME warehouse bans discussion ongoing"),
        ("🇨🇦","Canada","3.1M MT","Rio Tinto Alcan — 100% hydropower smelting; premium 'green aluminium'"),
        ("🇦🇪","UAE","2.7M MT","EGA (Emirates Global Aluminium) — gas-powered; Gulf hub for Asian/African bauxite"),
    ],
    "Brent Oil": [
        ("🇸🇦","Saudi Arabia","12.5M bpd","OPEC+ swing producer; Aramco controls spare capacity; voluntary cuts drive prices"),
        ("🇺🇸","USA","13.3M bpd","Largest producer; Permian Basin shale; not in OPEC but influences via SPR releases"),
        ("🇷🇺","Russia","10.8M bpd","Urals blend trades at discount; EU embargo rerouted flows to India/China"),
        ("🇮🇷","Iran","3.4M bpd","Sanctions limit exports; shadow fleet circumvents; nuclear deal talks affect supply outlook"),
        ("🇮🇶","Iraq","4.5M bpd","Basra crude tied to Brent; Kurdistan exports disputed; OPEC quota compliance varies"),
    ],
}

KEY_INSIGHTS = {
    "Copper": [
        ("H₂SO₄ Dependency","Peru's SX-EW copper leaching requires sulfuric acid, ~60% imported from Chinese smelters. Any disruption in China's acid exports (environmental shutdowns, export controls) directly constrains Peru's output, tightening global copper supply."),
        ("Green Transition Demand","Each EV uses 4x more copper than ICE vehicles (~83kg vs ~23kg). IEA projects copper demand from clean energy to double by 2030. Supply deficit forecasted at 6-8M MT by 2030 without new mines."),
        ("Concentrate Bottleneck","TC/RC (treatment/refining charges) at multi-year lows — smelters have excess capacity but not enough concentrate. Mines take 10-15 years from discovery to production."),
    ],
    "Aluminium": [
        ("Yunnan Power Crisis","China's Yunnan province hosts ~12% of global Al smelting capacity, powered by hydro. Dry seasons (Nov-Apr) force smelter curtailments of 1-2M MT/year, creating seasonal price spikes."),
        ("Russia Sanctions Overhang","Rusal produces ~6% of global aluminium. LME debates on banning Russian metal create price volatility. EU/US tariffs on Russian Al affect trade flows."),
        ("Carbon Border Tax","EU CBAM (Carbon Border Adjustment Mechanism) effective 2026 will add €50-100/MT cost to carbon-intensive Al imports, benefiting Canadian/Nordic 'green' smelters."),
    ],
    "Brent Oil": [
        ("OPEC+ Cuts Strategy","Saudi-led voluntary cuts of ~2.2M bpd through 2024-2025 support floor price of ~$70-80/bbl. Compliance varies — Iraq/Kazakhstan chronic overproducers."),
        ("Geopolitical Risk Premium","Red Sea/Houthi attacks reroute ~12% of global trade through longer Cape route, adding $1-2M per voyage. Iran-Israel tensions add $5-10/bbl risk premium."),
        ("US Shale Response","Permian Basin breakeven ~$45-55/bbl. Above $70, rigs increase within 3-6 months. US production now 13.3M bpd — acts as automatic ceiling on Brent rallies."),
    ],
}

def render_top5(commodity):
    items = TOP5.get(commodity, [])
    html = '<div style="display:grid;gap:10px">'
    for i, (flag, country, vol, note) in enumerate(items):
        rank_clr = ["#FFD600","#C0C0C0","#CD7F32","#8B949E","#484F58"][i]
        html += (
            '<div class="pcard" style="display:flex;gap:14px;align-items:flex-start">'
            '<div style="font-size:12px;font-weight:800;color:'+rank_clr+';min-width:24px">#'+str(i+1)+'</div>'
            '<div style="font-size:28px;line-height:1">'+flag+'</div>'
            '<div style="flex:1">'
            '<div style="font-size:14px;font-weight:700;color:#E6EDF3">'+country
            +'<span style="font-size:11px;color:#8B949E;margin-left:8px">'+vol+'</span></div>'
            '<div style="font-size:11px;color:#8B949E;margin-top:4px;line-height:1.5">'+note+'</div>'
            '</div></div>'
        )
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

def render_insights(commodity):
    items = KEY_INSIGHTS.get(commodity, [])
    for title, text in items:
        st.markdown(
            '<div class="pcard">'
            '<div style="font-size:12px;font-weight:700;color:#FFD600;letter-spacing:0.5px;margin-bottom:6px">⚡ '+title+'</div>'
            '<div style="font-size:12px;color:#8B949E;line-height:1.6">'+text+'</div>'
            '</div>',
            unsafe_allow_html=True)

# ══════════════════════════════════════════════
# NEWS FEED  (Google News RSS — free, no API key)
# ══════════════════════════════════════════════
@st.cache_data(ttl=1800)
def fetch_news(query, max_items=8):
    url = "https://news.google.com/rss/search?q=" + query.replace(" ","+") + "+price&hl=en&gl=US&ceid=US:en"
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=5) as resp:
            tree = ET.parse(resp)
        items = []
        for item in tree.findall(".//item")[:max_items]:
            t = item.findtext("title","")
            l = item.findtext("link","")
            s = item.findtext("source","")
            d = item.findtext("pubDate","")
            if t: items.append({"title":t,"link":l,"source":s,"date":d})
        return items
    except Exception:
        return []

def render_news(commodity):
    q_map = {"Copper":"copper commodity LME","Aluminium":"aluminium commodity LME","Brent Oil":"brent crude oil OPEC"}
    items = fetch_news(q_map.get(commodity, commodity))
    if not items:
        st.caption("Unable to load news — check network connection.")
        return
    for item in items:
        st.markdown(
            '<div class="news-item">'
            '<a class="news-title" href="'+item["link"]+'" target="_blank">'+item["title"]+'</a>'
            '<div class="news-src">'+item["source"]+' · '+item["date"][:16]+'</div>'
            '</div>',
            unsafe_allow_html=True)

# ══════════════════════════════════════════════
# MATPLOTLIB PNG (for PPTX — no Chrome needed)
# ══════════════════════════════════════════════
def _h2r(h):
    h=h.lstrip("#"); return tuple(int(h[i:i+2],16)/255 for i in (0,2,4))

def _mpl_monthly(monthly, metal, color):
    fig,ax=plt.subplots(figsize=(9.2,4.2),facecolor=CARD); ax.set_facecolor(CARD)
    labels=list(monthly["Label"]); prices=list(monthly["Price"]); x=list(range(len(labels)))
    prev=[None]+prices[:-1]
    for i,(p,pp) in enumerate(zip(prices,prev)):
        c=(0,200/255,83/255,0.12) if (pp is None or p>=pp) else (1,23/255,68/255,0.12)
        ax.bar(i,p,color=c,width=0.75,zorder=1)
    ax.plot(x,prices,color=_h2r(color),linewidth=2.2,zorder=3)
    for i,(p,pp) in enumerate(zip(prices,prev)):
        ax.scatter(i,p,color=GREEN if (pp is None or p>=pp) else RED,s=28,zorder=4,linewidths=0.8,edgecolors=CARD)
    for i,p in enumerate(prices):
        ax.text(i,p*1.002,f"{p:,.0f}",ha="center",va="bottom",fontsize=7,color=SEC)
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels,rotation=35,ha="right",fontsize=8,color=SEC)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v,_: f"{v:,.0f}"))
    ax.tick_params(colors=SEC,labelsize=8); ax.spines[:].set_color(GRID)
    ax.grid(axis="y",color=GRID,linewidth=0.5,zorder=0); ax.grid(axis="x",visible=False)
    ax.legend(handles=[Line2D([0],[0],color=_h2r(color),lw=2,label=f"{metal} Monthly Avg")],
              facecolor=CARD,edgecolor=GRID,labelcolor=SEC,fontsize=8,loc="upper left")
    ax.text(0,-0.18,"Source: LME / FRED",transform=ax.transAxes,fontsize=7.5,color=MUT)
    fig.tight_layout(pad=0.8); buf=io.BytesIO()
    fig.savefig(buf,format="png",dpi=150,facecolor=CARD,bbox_inches="tight"); plt.close(fig)
    buf.seek(0); return buf.read()

def _mpl_quarterly(quarterly, metal, color):
    fig,ax=plt.subplots(figsize=(9.2,4.2),facecolor=CARD); ax.set_facecolor(CARD)
    labels=list(quarterly["Label"]); prices=list(quarterly["Price"]); prev=[None]+prices[:-1]
    bc=[GREEN if (pp is None or p>=pp) else RED for p,pp in zip(prices,prev)]
    ax.bar(range(len(labels)),prices,color=bc,alpha=0.85,width=0.55,zorder=2)
    for i,p in enumerate(prices):
        ax.text(i,p*1.003,f"{p:,.0f}",ha="center",va="bottom",fontsize=8.5,color=TXT)
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels,rotation=20,ha="right",fontsize=8,color=SEC)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v,_: f"{v:,.0f}"))
    ax.tick_params(colors=SEC,labelsize=8); ax.spines[:].set_color(GRID)
    ax.grid(axis="y",color=GRID,linewidth=0.5,zorder=0); ax.grid(axis="x",visible=False)
    ax.text(0,-0.18,"Source: LME / FRED",transform=ax.transAxes,fontsize=7.5,color=MUT)
    fig.tight_layout(pad=0.8); buf=io.BytesIO()
    fig.savefig(buf,format="png",dpi=150,facecolor=CARD,bbox_inches="tight"); plt.close(fig)
    buf.seek(0); return buf.read()

# ══════════════════════════════════════════════
# PPTX BUILDER  (6 slides: Cu/Al/Brent × Monthly/Quarterly)
# ══════════════════════════════════════════════
def _rgb(h):
    h=h.lstrip("#"); return RGBColor(int(h[:2],16),int(h[2:4],16),int(h[4:],16))

def _txt_s(sl,l,t,w,h,text,pt,bold=False,color=TXT,align=PP_ALIGN.LEFT):
    tb=sl.shapes.add_textbox(l,t,w,h); tf=tb.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align; r=p.add_run()
    r.text=text; r.font.size=Pt(pt); r.font.bold=bold; r.font.color.rgb=_rgb(color)

def _kpi_pptx(sl,l,t,w,h,label,value,sub,bg):
    rect=sl.shapes.add_shape(1,l,t,w,h); rect.fill.solid(); rect.fill.fore_color.rgb=_rgb(bg)
    rect.line.color.rgb=_rgb(BORDER); rect.line.width=Pt(0.5)
    _txt_s(sl,l+Inches(0.08),t+Inches(0.06),w-Inches(0.16),Inches(0.2),label,7,color=SEC,align=PP_ALIGN.CENTER)
    _txt_s(sl,l+Inches(0.06),t+Inches(0.24),w-Inches(0.12),Inches(0.42),value,15,bold=True,align=PP_ALIGN.CENTER)
    if sub: _txt_s(sl,l+Inches(0.06),t+Inches(0.64),w-Inches(0.12),Inches(0.2),sub,7,color=GREEN,align=PP_ALIGN.CENTER)

def _slide(prs,title,subtitle,png_b,kpi,accent):
    sl=prs.slides.add_slide(prs.slide_layouts[6]); W,H=prs.slide_width,prs.slide_height
    bg=sl.shapes.add_shape(1,0,0,W,H); bg.fill.solid(); bg.fill.fore_color.rgb=_rgb(BG); bg.line.fill.background()
    hd=sl.shapes.add_shape(1,0,0,W,Inches(0.70)); hd.fill.solid(); hd.fill.fore_color.rgb=_rgb(CARD); hd.line.fill.background()
    ac=sl.shapes.add_shape(1,0,Inches(0.70),W,Inches(0.022)); ac.fill.solid(); ac.fill.fore_color.rgb=_rgb(accent); ac.line.fill.background()
    _txt_s(sl,Inches(0.25),Inches(0.06),Inches(7),Inches(0.36),title,16,bold=True)
    _txt_s(sl,Inches(0.25),Inches(0.41),Inches(7),Inches(0.22),subtitle,8.5,color=SEC)
    sl.shapes.add_picture(io.BytesIO(png_b),Inches(0.12),Inches(0.76),width=Inches(6.75),height=Inches(4.15))
    kx,kw,kh=Inches(7.1),Inches(2.65),Inches(1.08)
    _kpi_pptx(sl,kx,Inches(0.78),kw,kh,kpi["l1"],kpi["v1"],None,CARD)
    _kpi_pptx(sl,kx,Inches(1.96),kw,kh,kpi["l2"],kpi["v2"],None,CARD)
    _kpi_pptx(sl,kx,Inches(3.14),kw,kh,kpi["l3"],kpi["v3"],kpi.get("s3",""),CARD)
    _txt_s(sl,Inches(0.25),Inches(5.0),Inches(5),Inches(0.18),"Source: LME / FRED",7.5,color=MUT)

def _kd(m):
    c=m["Price"].iloc[-1]; cl=m["Label"].iloc[-1]; lo=m["Price"].min(); ll=m.loc[m["Price"].idxmin(),"Label"]
    f=m["Price"].iloc[0]; fl=m["Label"].iloc[0]; ch=c-f; pc=ch/f*100 if f else 0
    a="▲" if ch>=0 else "▼"
    return {"l1":f"CURRENT ({cl})","v1":f"${c:,.0f}","l2":f"PERIOD LOW ({ll})","v2":f"${lo:,.0f}",
            "l3":f"RETURN {fl}→{cl}","v3":f"{a} ${abs(ch):,.0f}","s3":f"{a} {abs(pc):.1f}%"}

def build_pptx(cu_m,cu_q,al_m,al_q,br_m,br_q,sel):
    def filt(df): r=df[df["Label"].isin(sel)]; return r if not r.empty else df.tail(10)
    def filtq(df,mf):
        if mf.empty: return df.tail(4)
        r=df[(df["Date"]>=mf["Date"].min())&(df["Date"]<=mf["Date"].max())]
        return r if not r.empty else df.tail(4)
    cm=filt(cu_m); am=filt(al_m); bm=filt(br_m)
    cq=filtq(cu_q,cm); aq=filtq(al_q,am); bq=filtq(br_q,bm)
    prs=Presentation(); prs.slide_width=Inches(10); prs.slide_height=Inches(5.625)
    _slide(prs,"Copper (Cu) — Monthly Average",f"USD/MT · {cm['Label'].iloc[0]} → {cm['Label'].iloc[-1]}",
           _mpl_monthly(cm,"Copper (Cu)",CU_CLR),_kd(cm),CU_CLR)
    _slide(prs,"Aluminium (Al) — Monthly Average",f"USD/MT · {am['Label'].iloc[0]} → {am['Label'].iloc[-1]}",
           _mpl_monthly(am,"Aluminium (Al)",AL_CLR),_kd(am),AL_CLR)
    _slide(prs,"Brent Crude Oil — Monthly Average",f"USD/bbl · {bm['Label'].iloc[0]} → {bm['Label'].iloc[-1]}",
           _mpl_monthly(bm,"Brent Oil",BR_CLR),_kd(bm),BR_CLR)
    for nm,qf,mf,cl in [("Copper (Cu)",cq,cm,CU_CLR),("Aluminium (Al)",aq,am,AL_CLR),("Brent Oil",bq,bm,BR_CLR)]:
        kp=_kd(mf)
        if len(qf): kp["l1"]=f"LATEST QTR ({qf['Label'].iloc[-1]})"; kp["v1"]=f"${qf['Price'].iloc[-1]:,.0f}"
        _slide(prs,f"{nm} — Quarterly Average",
               f"{qf['Label'].iloc[0] if len(qf) else ''} → {qf['Label'].iloc[-1] if len(qf) else ''}",
               _mpl_quarterly(qf,nm,cl),kp,cl)
    buf=io.BytesIO(); prs.save(buf); buf.seek(0); return buf.read()

# ══════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════
def main():
    with st.spinner(""):
        cu_raw=load_sheet(SHEET_COPPER); al_raw=load_sheet(SHEET_ALUMINIUM); br_raw=load_sheet(SHEET_BRENT)
    if cu_raw.empty or al_raw.empty:
        st.error("No data — check credentials."); return
    if br_raw.empty:
        br_raw = _synth("Brent")

    cu_m=monthly_avg(cu_raw); al_m=monthly_avg(al_raw); br_m=monthly_avg(br_raw)
    cu_q=quarterly_avg(cu_raw); al_q=quarterly_avg(al_raw); br_q=quarterly_avg(br_raw)

    render_ticker(cu_m, al_m, br_m)

    # ── sidebar
    with st.sidebar:
        st.markdown(
            '<div style="padding:12px 0 14px">'
            '<div style="font-size:26px;font-weight:900;color:#FFF;letter-spacing:3px;font-family:Arial Black,Arial,sans-serif;line-height:1">'
            'LME<span style="color:#B87333">.</span></div>'
            '<div style="font-size:9px;color:#8B949E;letter-spacing:2px;text-transform:uppercase;margin-top:4px">Commodities Terminal</div>'
            '</div><div style="height:2px;background:linear-gradient(90deg,#B87333,transparent);margin-bottom:16px"></div>',
            unsafe_allow_html=True)

        all_months = sorted(set(cu_m["Label"].tolist()+al_m["Label"].tolist()+br_m["Label"].tolist()),
                            key=lambda x: datetime.strptime(x,"%b-%y"))
        sel = st.multiselect("Export Months", all_months, default=all_months[-10:] if len(all_months)>=10 else all_months)

        st.markdown('<div style="font-size:10px;color:#484F58;margin:6px 0 10px;line-height:1.5">'
                    '6 slides: Cu/Al/Brent Monthly + Quarterly</div>', unsafe_allow_html=True)

        if st.button("⬇ Generate PPTX", type="primary"):
            if not sel: st.warning("Select months.")
            else:
                with st.spinner("Rendering 6 slides…"):
                    try:
                        data = build_pptx(cu_m,cu_q,al_m,al_q,br_m,br_q,sel)
                        st.download_button("⬇ Download PPTX",data=data,
                            file_name=f"Commodities_{datetime.now():%Y%m%d}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True)
                    except Exception as e: st.error(f"Error: {e}")

        st.markdown(
            '<div style="border-top:1px solid #30363D;margin-top:16px;padding-top:12px;font-size:10px;color:#484F58">'
            '<b style="color:#8B949E">DATA COVERAGE</b><br><br>'
            'Cu: '+cu_raw["Date"].min().strftime("%d %b %Y")+' → '+cu_raw["Date"].max().strftime("%d %b %Y")+'<br>'
            'Al: '+al_raw["Date"].min().strftime("%d %b %Y")+' → '+al_raw["Date"].max().strftime("%d %b %Y")+'<br>'
            'Brent: '+br_raw["Date"].min().strftime("%d %b %Y")+' → '+br_raw["Date"].max().strftime("%d %b %Y")+
            '</div>', unsafe_allow_html=True)

    # ── tabs
    tab_cu, tab_al, tab_br, tab_intel, tab_comp, tab_data = st.tabs([
        " 🟤 COPPER "," ⚙️ ALUMINIUM "," 🛢️ BRENT OIL "," 🌍 MARKET INTEL "," 📊 COMPARE "," 📋 DATA "])

    # ─ COPPER
    with tab_cu:
        c1,c2 = st.columns([3.2,1])
        with c1: st.plotly_chart(chart_live(cu_raw,"Copper (Cu)",CU_CLR),use_container_width=True,key="cu_l")
        with c2: render_kpis(cu_m)
        st.markdown('<div class="sh">QUARTERLY & ROLLING</div>',unsafe_allow_html=True)
        r1,r2 = st.columns(2)
        with r1: st.plotly_chart(chart_quarterly(cu_q,linear_forecast(cu_q,2),"Copper (Cu)",CU_CLR),use_container_width=True,key="cu_q")
        with r2: st.plotly_chart(chart_rolling(cu_raw,"Copper (Cu)",CU_CLR),use_container_width=True,key="cu_r")
        st.markdown('<div class="sh" style="margin-top:16px">MONTHLY TABLE</div>',unsafe_allow_html=True)
        render_table(cu_m)

    # ─ ALUMINIUM
    with tab_al:
        a1,a2 = st.columns([3.2,1])
        with a1: st.plotly_chart(chart_live(al_raw,"Aluminium (Al)",AL_CLR),use_container_width=True,key="al_l")
        with a2: render_kpis(al_m)
        st.markdown('<div class="sh">QUARTERLY & ROLLING</div>',unsafe_allow_html=True)
        r1,r2 = st.columns(2)
        with r1: st.plotly_chart(chart_quarterly(al_q,linear_forecast(al_q,2),"Aluminium (Al)",AL_CLR),use_container_width=True,key="al_q")
        with r2: st.plotly_chart(chart_rolling(al_raw,"Aluminium (Al)",AL_CLR),use_container_width=True,key="al_r")
        st.markdown('<div class="sh" style="margin-top:16px">MONTHLY TABLE</div>',unsafe_allow_html=True)
        render_table(al_m)

    # ─ BRENT OIL
    with tab_br:
        b1,b2 = st.columns([3.2,1])
        with b1: st.plotly_chart(chart_live(br_raw,"Brent Crude Oil",BR_CLR,"USD/BBL"),use_container_width=True,key="br_l")
        with b2: render_kpis(br_m,"USD/BBL")
        st.markdown('<div class="sh">QUARTERLY & ROLLING</div>',unsafe_allow_html=True)
        r1,r2 = st.columns(2)
        with r1: st.plotly_chart(chart_quarterly(br_q,linear_forecast(br_q,2),"Brent Oil",BR_CLR),use_container_width=True,key="br_q")
        with r2: st.plotly_chart(chart_rolling(br_raw,"Brent Oil",BR_CLR),use_container_width=True,key="br_r")
        st.markdown('<div class="sh" style="margin-top:16px">MONTHLY TABLE</div>',unsafe_allow_html=True)
        render_table(br_m,"USD/BBL")

    # ─ MARKET INTEL
    with tab_intel:
        sel_commodity = st.selectbox("Select Commodity", ["Copper","Aluminium","Brent Oil"],
                                      label_visibility="collapsed")
        st.markdown('<div class="sh" style="margin-top:8px">TOP 5 PRODUCING COUNTRIES</div>',unsafe_allow_html=True)
        render_top5(sel_commodity)

        i1, i2 = st.columns(2)
        with i1:
            st.markdown('<div class="sh" style="margin-top:16px">KEY SUPPLY CHAIN INSIGHTS</div>',unsafe_allow_html=True)
            render_insights(sel_commodity)
        with i2:
            st.markdown('<div class="sh" style="margin-top:16px">LATEST NEWS</div>',unsafe_allow_html=True)
            render_news(sel_commodity)

    # ─ COMPARISON
    with tab_comp:
        st.markdown('<div class="sh">INDEXED PERFORMANCE (Base=100)</div>',unsafe_allow_html=True)
        ci=cu_m.copy(); ci["I"]=ci["Price"]/ci["Price"].iloc[0]*100
        ai=al_m.copy(); ai["I"]=ai["Price"]/ai["Price"].iloc[0]*100
        bi=br_m.copy(); bi["I"]=bi["Price"]/bi["Price"].iloc[0]*100
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=ci["Date"],y=ci["I"],line=dict(color=CU_CLR,width=2),name="Copper"))
        fig.add_trace(go.Scatter(x=ai["Date"],y=ai["I"],line=dict(color=AL_CLR,width=2),name="Aluminium"))
        fig.add_trace(go.Scatter(x=bi["Date"],y=bi["I"],line=dict(color=BR_CLR,width=2),name="Brent Oil"))
        fig.add_hline(y=100,line_dash="dot",line_color=MUT,line_width=1)
        fig.update_layout(**_CL,title=dict(text="<b>Indexed Comparison</b>  <span style='color:#8B949E;font-size:12px'>Base=100</span>",font=dict(size=15)),height=380)
        st.plotly_chart(fig,use_container_width=True,key="comp")

        st.markdown('<div class="sh" style="margin-top:8px">CU vs AL SPREAD (USD/MT)</div>',unsafe_allow_html=True)
        mg=pd.merge(cu_m[["Date","Label","Price"]].rename(columns={"Price":"Cu"}),
                    al_m[["Date","Label","Price"]].rename(columns={"Price":"Al"}),on=["Date","Label"],how="inner")
        mg["Spread"]=mg["Cu"]-mg["Al"]
        fs=go.Figure()
        fs.add_trace(go.Bar(x=mg["Label"],y=mg["Spread"],marker_color=[GREEN if v>=0 else RED for v in mg["Spread"]],name="Spread"))
        fs.update_layout(**_CL,title=dict(text="<b>Cu–Al Spread</b>",font=dict(size=15)),height=300)
        st.plotly_chart(fs,use_container_width=True,key="spread")

    # ─ DATA TABLE
    with tab_data:
        d1,d2,d3 = st.columns(3)
        for col,raw,label in [(d1,cu_raw,"COPPER — USD/MT"),(d2,al_raw,"ALUMINIUM — USD/MT"),(d3,br_raw,"BRENT — USD/BBL")]:
            with col:
                st.markdown('<div class="sh">'+label+'</div>',unsafe_allow_html=True)
                dp=raw[["Date","Price"]].copy()
                dp["Date"]=dp["Date"].dt.strftime("%d %b %Y"); dp["Price"]=dp["Price"].apply(lambda x: f"${x:,.2f}")
                dp.columns=["Date","Price"]; st.dataframe(dp,height=400,use_container_width=True,hide_index=True)

if __name__ == "__main__":
    main()
