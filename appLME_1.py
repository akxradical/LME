"""
LME Metals Trading Dashboard
==============================
Stock-market style UI · Google Sheets live feed · 4-slide PPTX export

REQUIREMENTS (requirements.txt):
  streamlit>=1.32
  pandas
  numpy
  gspread
  google-auth
  plotly
  matplotlib
  python-pptx
  openpyxl

SECRETS (.streamlit/secrets.toml on Streamlit Cloud):
  SPREADSHEET_ID = "1zLyMANY56oFRwFug04WYavGUH_NAlRH8M3c-TXIRDlI"

  [gcp_service_account]
  type = "service_account"
  project_id = "lme-dashboard"
  private_key_id = "..."
  private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
  client_email = "lme-dashboard-materals@lme-dashboard.iam.gserviceaccount.com"
  client_id = "..."
  auth_uri = "https://accounts.google.com/o/oauth2/auth"
  token_uri = "https://oauth2.googleapis.com/token"
  auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
  client_x509_cert_url = "..."
  universe_domain = "googleapis.com"
"""

import io
from datetime import datetime

import gspread
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # headless — no display needed
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import FancyBboxPatch
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
SPREADSHEET_ID  = "1zLyMANY56oFRwFug04WYavGUH_NAlRH8M3c-TXIRDlI"
SHEET_COPPER    = "LME Copper"
SHEET_ALUMINIUM = "LME Aluminium"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Stock-market dark theme palette
BG_DARK    = "#0D1117"
BG_CARD    = "#161B22"
BG_CARD2   = "#1C2128"
BORDER     = "#30363D"
GREEN      = "#00C853"
GREEN_DIM  = "#1B5E20"
RED        = "#FF1744"
RED_DIM    = "#B71C1C"
BLUE       = "#2979FF"
GOLD       = "#FFD600"
COPPER_CLR = "#E87B35"
ALUM_CLR   = "#78909C"
TEXT_PRI   = "#E6EDF3"
TEXT_SEC   = "#8B949E"
TEXT_MUT   = "#484F58"
GRID       = "#21262D"

st.set_page_config(
    page_title="LME Metals Terminal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# GLOBAL CSS — dark terminal look
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
  /* ── page background ── */
  html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {{
      background-color: {BG_DARK} !important;
      color: {TEXT_PRI};
      font-family: 'JetBrains Mono', 'Fira Code', 'Courier New', monospace;
  }}
  [data-testid="stSidebar"] {{
      background-color: {BG_CARD} !important;
      border-right: 1px solid {BORDER};
  }}
  [data-testid="stHeader"] {{ background: transparent !important; }}
  .block-container {{ padding-top: 1rem; max-width: 100%; }}

  /* ── metric cards ── */
  .ticker-card {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 8px;
      padding: 14px 18px;
      margin-bottom: 10px;
  }}
  .ticker-label {{
      font-size: 10px;
      letter-spacing: 1.2px;
      color: {TEXT_SEC};
      text-transform: uppercase;
      margin-bottom: 4px;
  }}
  .ticker-value {{
      font-size: 28px;
      font-weight: 700;
      color: {TEXT_PRI};
      letter-spacing: -0.5px;
  }}
  .ticker-sub {{
      font-size: 11px;
      margin-top: 3px;
  }}
  .up   {{ color: {GREEN}; }}
  .down {{ color: {RED};   }}

  /* ── top ticker bar ── */
  .ticker-bar {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 8px;
      padding: 10px 20px;
      display: flex;
      gap: 40px;
      align-items: center;
      margin-bottom: 16px;
      flex-wrap: wrap;
  }}
  .ticker-item {{ display: flex; flex-direction: column; }}
  .ticker-name {{ font-size: 10px; color: {TEXT_SEC}; letter-spacing: 1px; }}
  .ticker-price {{ font-size: 18px; font-weight: 700; color: {TEXT_PRI}; }}
  .ticker-chg-up {{ font-size: 11px; color: {GREEN}; }}
  .ticker-chg-dn {{ font-size: 11px; color: {RED};   }}

  /* ── section headers ── */
  .section-header {{
      font-size: 11px;
      letter-spacing: 1.5px;
      color: {TEXT_SEC};
      text-transform: uppercase;
      border-bottom: 1px solid {BORDER};
      padding-bottom: 6px;
      margin-bottom: 12px;
  }}

  /* ── data table ── */
  .data-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
  }}
  .data-table th {{
      background: {BG_CARD2};
      color: {TEXT_SEC};
      font-size: 10px;
      letter-spacing: 1px;
      text-transform: uppercase;
      padding: 8px 12px;
      text-align: right;
      border-bottom: 1px solid {BORDER};
  }}
  .data-table th:first-child {{ text-align: left; }}
  .data-table td {{
      padding: 7px 12px;
      border-bottom: 1px solid {BG_CARD2};
      text-align: right;
      color: {TEXT_PRI};
  }}
  .data-table td:first-child {{ text-align: left; color: {TEXT_SEC}; }}
  .data-table tr:hover td {{ background: {BG_CARD2}; }}

  /* ── tabs ── */
  .stTabs [data-baseweb="tab-list"] {{
      background: {BG_CARD};
      border-radius: 8px;
      padding: 4px;
      gap: 4px;
      border: 1px solid {BORDER};
  }}
  .stTabs [data-baseweb="tab"] {{
      background: transparent;
      color: {TEXT_SEC};
      border-radius: 6px;
      padding: 8px 20px;
      font-size: 12px;
      letter-spacing: 0.5px;
  }}
  .stTabs [aria-selected="true"] {{
      background: {BG_CARD2} !important;
      color: {TEXT_PRI} !important;
      border: 1px solid {BORDER} !important;
  }}

  /* ── sidebar widgets ── */
  .stMultiSelect > div, .stSelectbox > div {{
      background: {BG_CARD2};
      border: 1px solid {BORDER};
      border-radius: 6px;
  }}
  .stButton > button {{
      background: {BLUE};
      color: white;
      border: none;
      border-radius: 6px;
      font-weight: 600;
      letter-spacing: 0.5px;
      width: 100%;
      padding: 10px;
  }}
  .stButton > button:hover {{ background: #1565C0; }}
  [data-testid="stDownloadButton"] > button {{
      background: {GREEN_DIM};
      color: {GREEN};
      border: 1px solid {GREEN};
      border-radius: 6px;
      font-weight: 600;
      width: 100%;
      padding: 10px;
  }}

  /* hide streamlit branding */
  #MainMenu, footer, header {{ visibility: hidden; }}
  [data-testid="stToolbar"] {{ display: none; }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_sheet(sheet_name: str) -> pd.DataFrame:
    try:
        info  = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        gc    = gspread.authorize(creds)
        sid   = st.secrets.get("SPREADSHEET_ID", SPREADSHEET_ID)
        ws    = gc.open_by_key(sid).worksheet(sheet_name)
        rows  = ws.get_all_records()
        df    = pd.DataFrame(rows)
        df.columns = [c.strip() for c in df.columns]
        date_col  = next(c for c in df.columns if "date"  in c.lower())
        price_col = next(c for c in df.columns
                         if "price" in c.lower() or "cash" in c.lower())
        df = df.rename(columns={date_col: "Date", price_col: "Price"})
        df["Date"]  = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        df["Price"] = pd.to_numeric(
            df["Price"].astype(str).str.replace(",", ""), errors="coerce"
        )
        return (df.dropna(subset=["Date", "Price"])
                  .sort_values("Date")
                  .reset_index(drop=True))
    except Exception:
        return _synthetic(sheet_name)


def _synthetic(sheet_name: str) -> pd.DataFrame:
    rng   = np.random.default_rng(42 if "Copper" in sheet_name else 7)
    base  = 9000 if "Copper" in sheet_name else 2300
    dates = pd.date_range("2025-01-02", datetime.today(), freq="B")
    noise = rng.normal(0, 60, len(dates)).cumsum()
    prices= np.clip(base + noise, base * 0.8, base * 1.35)
    return pd.DataFrame({"Date": dates, "Price": prices.round(2)})

# ─────────────────────────────────────────────
# ANALYTICS
# ─────────────────────────────────────────────
def monthly_avg(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["YM"]   = df["Date"].dt.to_period("M")
    m          = df.groupby("YM")["Price"].mean().reset_index()
    m["Date"]  = m["YM"].dt.to_timestamp()
    m["Label"] = m["YM"].dt.strftime("%b-%y")
    m["Price"] = m["Price"].round(2)
    return m.drop(columns="YM")

def quarterly_avg(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["YQ"]   = df["Date"].dt.to_period("Q")
    q          = df.groupby("YQ")["Price"].mean().reset_index()
    q["Date"]  = q["YQ"].dt.to_timestamp()
    q["Label"] = q["YQ"].dt.strftime("Q%q-%Y")
    q["Price"] = q["Price"].round(2)
    return q.drop(columns="YQ")

def rolling_avg(df: pd.DataFrame, w=3) -> pd.DataFrame:
    d = df.copy().sort_values("Date")
    d[f"MA{w}"] = d["Price"].rolling(w).mean().round(2)
    return d

def linear_forecast(monthly: pd.DataFrame, periods=3) -> pd.DataFrame:
    if len(monthly) < 2:
        return pd.DataFrame(columns=["Date", "Price", "Label"])
    x = np.arange(len(monthly))
    y = monthly["Price"].values
    s, i = np.polyfit(x, y, 1)
    last = monthly["Date"].iloc[-1]
    fd = [last + pd.DateOffset(months=k+1) for k in range(periods)]
    fp = [i + s*(len(monthly)+k+1) for k in range(periods)]
    return pd.DataFrame({
        "Date":  fd,
        "Price": np.round(fp, 2),
        "Label": [d.strftime("%b-%y") for d in fd],
    })

def mom_change(monthly: pd.DataFrame) -> pd.DataFrame:
    m = monthly.copy()
    m["Change"]   = m["Price"].diff().round(2)
    m["Change%"]  = (m["Price"].pct_change() * 100).round(2)
    m["3M_Avg"]   = m["Price"].rolling(3).mean().round(2)
    return m

# ─────────────────────────────────────────────
# CHART BUILDERS  (dark theme)
# ─────────────────────────────────────────────
_CHART_LAYOUT = dict(
    paper_bgcolor=BG_CARD,
    plot_bgcolor=BG_CARD,
    font=dict(color=TEXT_PRI, family="'JetBrains Mono','Fira Code',monospace", size=11),
    margin=dict(l=55, r=15, t=65, b=55),
    xaxis=dict(showgrid=False, zeroline=False,
               color=TEXT_SEC, linecolor=BORDER, tickcolor=BORDER),
    yaxis=dict(gridcolor=GRID, zeroline=False,
               color=TEXT_SEC, linecolor=BORDER, tickformat=","),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor=BORDER,
                orientation="h", y=-0.18, font=dict(size=10)),
)

def candlestick_monthly(monthly: pd.DataFrame, forecast: pd.DataFrame,
                        metal: str, color: str) -> go.Figure:
    """OHLC-style bar chart: monthly avg with up/down coloring."""
    monthly = monthly.copy()
    monthly["prev"]    = monthly["Price"].shift(1)
    monthly["is_up"]   = monthly["Price"] >= monthly["prev"]
    colors_bar         = [GREEN if u else RED for u in monthly["is_up"]]

    fig = go.Figure()

    # Volume-style background bars  (rgba — Plotly does NOT accept hex+alpha)
    GREEN_DIM_BAR = "rgba(0,200,83,0.15)"
    RED_DIM_BAR   = "rgba(255,23,68,0.15)"
    fig.add_trace(go.Bar(
        x=monthly["Label"], y=monthly["Price"],
        marker_color=[GREEN_DIM_BAR if u else RED_DIM_BAR for u in monthly["is_up"]],
        showlegend=False, hoverinfo="skip",
    ))

    # Main line
    fig.add_trace(go.Scatter(
        x=monthly["Label"], y=monthly["Price"],
        mode="lines+markers+text",
        line=dict(color=color, width=2),
        marker=dict(size=5, color=colors_bar, line=dict(width=1, color=BG_DARK)),
        text=[f"{p:,.0f}" for p in monthly["Price"]],
        textposition="top center",
        textfont=dict(size=8, color=TEXT_SEC),
        name=f"{metal} Monthly Avg",
    ))

    # Forecast
    if len(forecast):
        bx = [monthly["Label"].iloc[-1]] + list(forecast["Label"])
        by = [monthly["Price"].iloc[-1]] + list(forecast["Price"])
        fig.add_trace(go.Scatter(
            x=bx, y=by,
            mode="lines+markers+text",
            line=dict(color=GOLD, width=1.5, dash="dot"),
            marker=dict(size=6, symbol="diamond", color=GOLD),
            text=[""] + [f"{p:,.0f}" for p in forecast["Price"]],
            textposition="top center",
            textfont=dict(size=8, color=GOLD),
            name="Forecast ▸",
        ))

    first_lbl = monthly["Label"].iloc[0]
    last_lbl  = monthly["Label"].iloc[-1]
    fig.update_layout(
        **_CHART_LAYOUT,
        title=dict(
            text=f"<b>{metal}</b>  <span style='color:{TEXT_SEC};font-size:12px'>"
                 f"Monthly Avg  ·  {first_lbl} → {last_lbl}</span>",
            font=dict(size=15, color=TEXT_PRI),
        ),
        height=380,
        barmode="overlay",
        annotations=[dict(text="Source: LME", x=0, y=-0.18, xref="paper",
                          yref="paper", showarrow=False,
                          font=dict(size=9, color=TEXT_MUT))],
    )
    return fig


def area_chart(monthly: pd.DataFrame, forecast: pd.DataFrame,
               metal: str, color: str) -> go.Figure:
    """Area chart with gradient fill — TradingView style."""
    fig = go.Figure()
    r, g, b = bytes.fromhex(color.lstrip("#"))

    fig.add_trace(go.Scatter(
        x=monthly["Date"], y=monthly["Price"],
        fill="tozeroy",
        fillcolor=f"rgba({r},{g},{b},0.10)",
        line=dict(color=color, width=2),
        mode="lines",
        name=f"{metal}",
    ))
    # Dots at each month end
    fig.add_trace(go.Scatter(
        x=monthly["Date"], y=monthly["Price"],
        mode="markers+text",
        marker=dict(size=5, color=color,
                    line=dict(width=1, color=BG_DARK)),
        text=[f"{p:,.0f}" for p in monthly["Price"]],
        textposition="top center",
        textfont=dict(size=8, color=TEXT_SEC),
        showlegend=False,
    ))

    if len(forecast):
        bx = [monthly["Date"].iloc[-1]] + list(forecast["Date"])
        by = [monthly["Price"].iloc[-1]] + list(forecast["Price"])
        fig.add_trace(go.Scatter(
            x=bx, y=by,
            fill="tozeroy",
            fillcolor=f"rgba(255,214,0,0.05)",
            line=dict(color=GOLD, width=1.5, dash="dot"),
            mode="lines+markers",
            marker=dict(size=6, symbol="diamond", color=GOLD),
            name="Forecast ▸",
        ))

    fig.update_layout(
        **_CHART_LAYOUT,
        title=dict(
            text=f"<b>{metal}</b>  <span style='color:{TEXT_SEC};font-size:12px'>"
                 f"Price Trend</span>",
            font=dict(size=15, color=TEXT_PRI),
        ),
        height=350,
    )
    return fig


def quarterly_chart(quarterly: pd.DataFrame, forecast_q: pd.DataFrame,
                    metal: str, color: str) -> go.Figure:
    quarterly = quarterly.copy()
    quarterly["prev"]  = quarterly["Price"].shift(1)
    quarterly["is_up"] = quarterly["Price"] >= quarterly["prev"]
    bar_colors = [GREEN if u else RED for u in quarterly["is_up"]]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=quarterly["Label"], y=quarterly["Price"],
        marker_color=bar_colors,
        marker_line_width=0,
        opacity=0.85,
        text=[f"{p:,.0f}" for p in quarterly["Price"]],
        textposition="outside",
        textfont=dict(size=9, color=TEXT_PRI),
        name="Quarterly Avg",
    ))
    if len(forecast_q):
        fig.add_trace(go.Bar(
            x=forecast_q["Label"], y=forecast_q["Price"],
            marker_color=GOLD,
            marker_line_width=0,
            opacity=0.5,
            text=[f"{p:,.0f}" for p in forecast_q["Price"]],
            textposition="outside",
            textfont=dict(size=9, color=GOLD),
            name="Forecast ▸",
        ))
    fig.update_layout(
        **_CHART_LAYOUT,
        title=dict(
            text=f"<b>{metal}</b>  <span style='color:{TEXT_SEC};font-size:12px'>"
                 f"Quarterly Average</span>",
            font=dict(size=15, color=TEXT_PRI),
        ),
        height=350,
        barmode="group",
    )
    return fig


def rolling_chart(df_raw: pd.DataFrame, metal: str, color: str) -> go.Figure:
    r3 = rolling_avg(df_raw, 3)
    r6 = rolling_avg(df_raw, 6)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=r3["Date"], y=r3["Price"],
        line=dict(color=BORDER, width=0.8),
        mode="lines", name="Daily", opacity=0.5,
    ))
    fig.add_trace(go.Scatter(
        x=r3["Date"], y=r3["MA3"],
        line=dict(color=color, width=2),
        mode="lines", name="3M MA",
    ))
    fig.add_trace(go.Scatter(
        x=r6["Date"], y=r6["MA6"],
        line=dict(color=GOLD, width=1.5, dash="dash"),
        mode="lines", name="6M MA",
    ))
    fig.update_layout(
        **_CHART_LAYOUT,
        title=dict(
            text=f"<b>{metal}</b>  <span style='color:{TEXT_SEC};font-size:12px'>"
                 f"Rolling Averages (3M / 6M)</span>",
            font=dict(size=15, color=TEXT_PRI),
        ),
        height=350,
    )
    return fig


def chart_live_daily(df_raw: pd.DataFrame, metal: str, color: str) -> go.Figure:
    """Main live chart — shows ALL daily prices like a stock chart (TradingView style)."""
    df = df_raw.copy().sort_values("Date").reset_index(drop=True)
    r3 = rolling_avg(df, 3)

    # Compute daily colour: green if price >= prev, red otherwise
    df["prev"]   = df["Price"].shift(1)
    df["is_up"]  = df["Price"] >= df["prev"]

    cur   = df["Price"].iloc[-1]
    prev  = df["Price"].iloc[-2] if len(df) > 1 else cur
    chg   = cur - prev
    pct   = chg / prev * 100 if prev else 0
    arrow = "▲" if chg >= 0 else "▼"
    clr   = GREEN if chg >= 0 else RED
    r,g,b = bytes.fromhex(color.lstrip("#"))

    fig = go.Figure()

    # Area fill under daily price
    fig.add_trace(go.Scatter(
        x=df["Date"], y=df["Price"],
        fill="tozeroy",
        fillcolor=f"rgba({r},{g},{b},0.07)",
        line=dict(color=color, width=1.5),
        mode="lines",
        name="Daily Price",
        hovertemplate="<b>%{x|%d %b %Y}</b><br>$%{y:,.2f} /MT<extra></extra>",
    ))

    # 3M moving average overlay
    fig.add_trace(go.Scatter(
        x=r3["Date"], y=r3["MA3"],
        line=dict(color=GOLD, width=1.2, dash="dot"),
        mode="lines",
        name="3M Moving Avg",
        hovertemplate="3M Avg: $%{y:,.2f}<extra></extra>",
    ))

    first_lbl = df["Date"].iloc[0].strftime("%d %b %Y")
    last_lbl  = df["Date"].iloc[-1].strftime("%d %b %Y")

    # Build layout without xaxis/yaxis from _CHART_LAYOUT to avoid duplicate kwarg error
    _base = {k: v for k, v in _CHART_LAYOUT.items() if k not in ("xaxis", "yaxis")}
    fig.update_layout(
        **_base,
        title=dict(
            text=(f"<b>{metal}</b>  "
                  f"<span style='font-size:22px;font-weight:700;color:{clr}'>"
                  f"${cur:,.2f}</span>  "
                  f"<span style='font-size:13px;color:{clr}'>{arrow} ${abs(chg):,.2f}  ({arrow}{abs(pct):.2f}%)</span>"
                  f"<br><span style='font-size:10px;color:{TEXT_SEC}'>"
                  f"USD/MT  ·  {first_lbl} \u2192 {last_lbl}  ·  LME Cash Settlement</span>"),
            font=dict(size=15, color=TEXT_PRI),
        ),
        xaxis=dict(
            showgrid=False, zeroline=False,
            color=TEXT_SEC, linecolor=BORDER, tickcolor=BORDER,
            rangeslider=dict(visible=True, bgcolor=BG_DARK, thickness=0.04),
            rangeselector=dict(
                bgcolor=BG_CARD2, activecolor=BORDER,
                font=dict(color=TEXT_SEC, size=10),
                buttons=[
                    dict(count=1, label="1M",  step="month", stepmode="backward"),
                    dict(count=3, label="3M",  step="month", stepmode="backward"),
                    dict(count=6, label="6M",  step="month", stepmode="backward"),
                    dict(count=1, label="YTD", step="year",  stepmode="todate"),
                    dict(step="all", label="ALL"),
                ],
            ),
        ),
        yaxis=dict(gridcolor=GRID, zeroline=False, color=TEXT_SEC,
                   linecolor=BORDER, tickformat=",", side="right"),
        height=420,
        hovermode="x unified",
        annotations=[dict(
            text="Source: LME via Google Sheets",
            x=0, y=-0.08, xref="paper", yref="paper",
            showarrow=False, font=dict(size=9, color=TEXT_MUT),
        )],
    )
    return fig

# ─────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────
def render_kpis(monthly: pd.DataFrame):
    cur      = monthly["Price"].iloc[-1]
    cur_lbl  = monthly["Label"].iloc[-1]
    low_val  = monthly["Price"].min()
    low_lbl  = monthly.loc[monthly["Price"].idxmin(), "Label"]
    high_val = monthly["Price"].max()
    high_lbl = monthly.loc[monthly["Price"].idxmax(), "Label"]
    prev     = monthly["Price"].iloc[-2] if len(monthly) > 1 else cur
    chg      = cur - prev
    pct      = chg / prev * 100 if prev else 0
    first    = monthly["Price"].iloc[0]
    ytd_chg  = cur - first
    ytd_pct  = ytd_chg / first * 100 if first else 0

    up_dn      = "up"   if chg     >= 0 else "down"
    up_dn_ytd  = "up"   if ytd_chg >= 0 else "down"
    arrow      = "▲"    if chg     >= 0 else "▼"
    arrow_ytd  = "▲"    if ytd_chg >= 0 else "▼"

    st.markdown(f"""
    <div class="ticker-card">
      <div class="ticker-label">CURRENT · {cur_lbl}</div>
      <div class="ticker-value">${cur:,.2f}</div>
      <div class="ticker-sub {up_dn}">{arrow} ${abs(chg):,.2f} &nbsp;({arrow} {abs(pct):.2f}%) MoM</div>
    </div>

    <div class="ticker-card">
      <div class="ticker-label">PERIOD HIGH · {high_lbl}</div>
      <div class="ticker-value up">${high_val:,.2f}</div>
      <div class="ticker-sub" style="color:{TEXT_MUT}">USD / Metric Tonne</div>
    </div>

    <div class="ticker-card">
      <div class="ticker-label">PERIOD LOW · {low_lbl}</div>
      <div class="ticker-value down">${low_val:,.2f}</div>
      <div class="ticker-sub" style="color:{TEXT_MUT}">USD / Metric Tonne</div>
    </div>

    <div class="ticker-card">
      <div class="ticker-label">YTD RETURN</div>
      <div class="ticker-value {up_dn_ytd}">{arrow_ytd} {abs(ytd_pct):.1f}%</div>
      <div class="ticker-sub {up_dn_ytd}">{arrow_ytd} ${abs(ytd_chg):,.0f} /MT since {monthly["Label"].iloc[0]}</div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# MONTHLY TABLE  (no matplotlib)
# ─────────────────────────────────────────────
def render_monthly_table(monthly: pd.DataFrame, color: str):
    m = mom_change(monthly).copy()
    rows = ""
    for _, row in m.iterrows():
        chg_html = ""
        if not pd.isna(row["Change"]):
            cls  = "up" if row["Change"] >= 0 else "down"
            sign = "+" if row["Change"] >= 0 else ""
            chg_html = f'<span class="{cls}">{sign}{row["Change"]:,.2f}</span>'
        pct_html = ""
        if not pd.isna(row["Change%"]):
            cls  = "up" if row["Change%"] >= 0 else "down"
            sign = "+" if row["Change%"] >= 0 else ""
            pct_html = f'<span class="{cls}">{sign}{row["Change%"]:.2f}%</span>'
        avg3 = f'${row["3M_Avg"]:,.2f}' if not pd.isna(row["3M_Avg"]) else "—"
        rows += f"""
        <tr>
          <td>{row["Label"]}</td>
          <td>${row["Price"]:,.2f}</td>
          <td>{chg_html}</td>
          <td>{pct_html}</td>
          <td style="color:{TEXT_SEC}">{avg3}</td>
        </tr>"""
    st.markdown(f"""
    <table class="data-table">
      <thead>
        <tr>
          <th>Month</th>
          <th>Avg Price (USD/MT)</th>
          <th>MoM Change</th>
          <th>MoM %</th>
          <th>3M Avg</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TOP TICKER BAR
# ─────────────────────────────────────────────
def render_ticker_bar(cu_m: pd.DataFrame, al_m: pd.DataFrame):
    def _stats(df):
        cur  = df["Price"].iloc[-1]
        prev = df["Price"].iloc[-2] if len(df) > 1 else cur
        chg  = cur - prev
        pct  = chg / prev * 100 if prev else 0
        return cur, chg, pct

    cu_cur, cu_chg, cu_pct = _stats(cu_m)
    al_cur, al_chg, al_pct = _stats(al_m)

    cu_badge   = 'lme-cup' if cu_chg >= 0 else 'lme-cdn'
    al_badge   = 'lme-cup' if al_chg >= 0 else 'lme-cdn'
    cu_arr     = '&#9650;' if cu_chg >= 0 else '&#9660;'
    al_arr     = '&#9650;' if al_chg >= 0 else '&#9660;'
    cu_price_s = '${:,.2f}'.format(cu_cur)
    al_price_s = '${:,.2f}'.format(al_cur)
    cu_pct_s   = '{} {:.2f}%'.format(cu_arr, abs(cu_pct))
    al_pct_s   = '{} {:.2f}%'.format(al_arr, abs(al_pct))
    cu_abs_s   = '{} ${:,.2f} MoM'.format(cu_arr, abs(cu_chg))
    al_abs_s   = '{} ${:,.2f} MoM'.format(al_arr, abs(al_chg))
    now        = datetime.now().strftime('%d %b %Y  %H:%M')

    css = (
        '<style>'
        '.lme-wrap{background:#0A0F1E;border-bottom:2px solid #B87333;margin-bottom:12px}'
        '.lme-top{background:#050A14;padding:7px 24px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #1C2333}'
        '.lme-logo{font-size:22px;font-weight:900;letter-spacing:3px;color:#FFF;font-family:"Arial Black",Arial,sans-serif}'
        '.lme-dot-o{color:#B87333}'
        '.lme-sub{font-size:9px;color:#8B949E;letter-spacing:2px;text-transform:uppercase;margin-top:1px}'
        '.lme-bar{background:#0A0F1E;padding:14px 28px;display:flex;align-items:stretch;border-bottom:1px solid #1C2333}'
        '.lme-pb{display:flex;flex-direction:column;padding:8px 36px 8px 0;margin-right:32px;border-right:1px solid #1C2333;min-width:210px}'
        '.lme-mn{font-size:10px;letter-spacing:1.5px;color:#8B949E;text-transform:uppercase;margin-bottom:5px;font-family:Arial,sans-serif}'
        '.lme-pr{display:flex;align-items:baseline;gap:10px}'
        '.lme-pv{font-size:27px;font-weight:700;color:#E6EDF3;letter-spacing:-0.5px;font-family:Arial,sans-serif}'
        '.lme-pu{font-size:11px;color:#484F58;font-family:Arial,sans-serif}'
        '.lme-cr{display:flex;align-items:center;gap:8px;margin-top:4px}'
        '.lme-badge{font-size:12px;font-weight:600;padding:2px 9px;border-radius:3px;font-family:Arial,sans-serif}'
        '.lme-cup{background:rgba(0,200,83,0.15);color:#00C853}'
        '.lme-cdn{background:rgba(255,23,68,0.15);color:#FF1744}'
        '.lme-ab{font-size:11px;color:#8B949E;font-family:Arial,sans-serif}'
        '.lme-rf{margin-left:auto;display:flex;flex-direction:column;justify-content:center;padding-left:28px;border-left:1px solid #1C2333}'
        '.lme-live{display:inline-block;width:7px;height:7px;background:#00C853;border-radius:50%;margin-right:5px;vertical-align:middle;animation:ldp 2s infinite}'
        '@keyframes ldp{0%,100%{opacity:1}50%{opacity:0.3}}'
        '</style>'
    )

    body = (
        '<div class="lme-wrap">'
        '<div class="lme-top">'
        '<div>'
        '<div class="lme-logo">LME<span class="lme-dot-o">.</span></div>'
        '<div class="lme-sub">London Metal Exchange &middot; Metals Price Terminal</div>'
        '</div>'
        '<div style="font-size:10px;color:#484F58;letter-spacing:1px">CASH SETTLEMENT &middot; USD/MT</div>'
        '</div>'
        '<div class="lme-bar">'
        '<div class="lme-pb">'
        '<div class="lme-mn">&#9632; Copper (Cu)</div>'
        '<div class="lme-pr"><span class="lme-pv">' + cu_price_s + '</span><span class="lme-pu">USD/MT</span></div>'
        '<div class="lme-cr"><span class="lme-badge ' + cu_badge + '">' + cu_pct_s + '</span><span class="lme-ab">' + cu_abs_s + '</span></div>'
        '</div>'
        '<div class="lme-pb">'
        '<div class="lme-mn">&#9632; Aluminium (Al)</div>'
        '<div class="lme-pr"><span class="lme-pv">' + al_price_s + '</span><span class="lme-pu">USD/MT</span></div>'
        '<div class="lme-cr"><span class="lme-badge ' + al_badge + '">' + al_pct_s + '</span><span class="lme-ab">' + al_abs_s + '</span></div>'
        '</div>'
        '<div class="lme-rf">'
        '<span style="font-size:9px;color:#484F58;letter-spacing:1.5px;text-transform:uppercase"><span class="lme-live"></span>Live Feed</span>'
        '<span style="font-size:14px;color:#8B949E;margin-top:3px;font-family:Arial,sans-serif">' + now + '</span>'
        '<span style="font-size:9px;color:#484F58;margin-top:2px">Auto-refresh every 5 min</span>'
        '</div>'
        '</div>'
        '</div><br>'
    )
    st.markdown(css + body, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PPTX BUILDER
# ─────────────────────────────────────────────
def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

def _txt(slide, l, t, w, h, text, pt, bold=False,
         color="#E6EDF3", align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p  = tf.paragraphs[0]
    p.alignment = align
    r  = p.add_run()
    r.text = text
    r.font.size  = Pt(pt)
    r.font.bold  = bold
    r.font.color.rgb = _rgb(color)

def _kpi_card_pptx(slide, left, top, w, h, label, value, sub, bg):
    rect = slide.shapes.add_shape(1, left, top, w, h)
    rect.fill.solid(); rect.fill.fore_color.rgb = _rgb(bg)
    rect.line.color.rgb = _rgb(BORDER)
    rect.line.width = Pt(0.5)
    _txt(slide, left+Inches(0.08), top+Inches(0.06), w-Inches(0.16), Inches(0.2),
         label, 7, color="#8B949E", align=PP_ALIGN.CENTER)
    _txt(slide, left+Inches(0.06), top+Inches(0.24), w-Inches(0.12), Inches(0.42),
         value, 15, bold=True, color=TEXT_PRI, align=PP_ALIGN.CENTER)
    if sub:
        _txt(slide, left+Inches(0.06), top+Inches(0.64), w-Inches(0.12), Inches(0.2),
             sub, 7, color=GREEN, align=PP_ALIGN.CENTER)

def _build_slide(prs, title, subtitle, chart_png, kpi_dict, accent):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H  = prs.slide_width, prs.slide_height

    # Dark background
    bg = slide.shapes.add_shape(1, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = _rgb(BG_DARK)
    bg.line.fill.background()

    # Header bar
    hd = slide.shapes.add_shape(1, 0, 0, W, Inches(0.70))
    hd.fill.solid(); hd.fill.fore_color.rgb = _rgb(BG_CARD)
    hd.line.fill.background()

    # Accent stripe
    ac = slide.shapes.add_shape(1, 0, Inches(0.70), W, Inches(0.022))
    ac.fill.solid(); ac.fill.fore_color.rgb = _rgb(accent)
    ac.line.fill.background()

    # Title + subtitle
    _txt(slide, Inches(0.25), Inches(0.06), Inches(7), Inches(0.36),
         title, 16, bold=True, color=TEXT_PRI)
    _txt(slide, Inches(0.25), Inches(0.41), Inches(7), Inches(0.22),
         subtitle, 8.5, color=TEXT_SEC)

    # Chart image
    slide.shapes.add_picture(
        io.BytesIO(chart_png),
        Inches(0.12), Inches(0.76),
        width=Inches(6.75), height=Inches(4.15),
    )

    # KPI cards
    kx, kw, kh = Inches(7.1), Inches(2.65), Inches(1.08)
    _kpi_card_pptx(slide, kx, Inches(0.78), kw, kh,
                   kpi_dict["l1"], kpi_dict["v1"], None, BG_CARD)
    _kpi_card_pptx(slide, kx, Inches(1.96), kw, kh,
                   kpi_dict["l2"], kpi_dict["v2"], None, BG_CARD)
    _kpi_card_pptx(slide, kx, Inches(3.14), kw, kh,
                   kpi_dict["l3"], kpi_dict["v3"], kpi_dict.get("s3",""), BG_CARD)

    # Footer
    _txt(slide, Inches(0.25), Inches(5.0), Inches(5), Inches(0.18),
         "Source: LME  ·  USD per Metric Tonne", 7.5, color=TEXT_MUT)

def _kpi_dict(monthly):
    cur  = monthly["Price"].iloc[-1]; cl = monthly["Label"].iloc[-1]
    lo   = monthly["Price"].min();   ll = monthly.loc[monthly["Price"].idxmin(),"Label"]
    hi   = monthly["Price"].max();   hl = monthly.loc[monthly["Price"].idxmax(),"Label"]
    f    = monthly["Price"].iloc[0]; fl = monthly["Label"].iloc[0]
    chg  = cur - f; pct = chg/f*100 if f else 0
    a    = "▲" if chg >= 0 else "▼"
    return {
        "l1": f"CURRENT ({cl})", "v1": f"${cur:,.0f} /MT",
        "l2": f"PERIOD LOW ({ll})", "v2": f"${lo:,.0f} /MT",
        "l3": f"RETURN  {fl}→{cl}",
        "v3": f"{a} ${abs(chg):,.0f} /MT",
        "s3": f"{a} {abs(pct):.1f}%",
    }

# ─────────────────────────────────────────────
# MATPLOTLIB CHART → PNG  (no Chrome / kaleido)
# ─────────────────────────────────────────────
def _hex_to_rgb01(h: str):
    """Convert '#RRGGBB' → (r, g, b) floats 0-1."""
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))


def _mpl_monthly_png(monthly: pd.DataFrame, forecast: pd.DataFrame,
                     metal: str, color: str) -> bytes:
    """Dark-themed monthly line chart rendered by matplotlib — returns PNG bytes."""
    BG   = "#0D1117"
    CARD = "#161B22"
    GRD  = "#21262D"
    SEC  = "#8B949E"
    PRI  = "#E6EDF3"
    G    = "#00C853"
    R    = "#FF1744"
    GOLD_C = "#FFD600"

    fig, ax = plt.subplots(figsize=(9.2, 4.2), facecolor=CARD)
    ax.set_facecolor(CARD)

    labels = list(monthly["Label"])
    prices = list(monthly["Price"])
    x_idx  = list(range(len(labels)))

    # Up / down bar fill
    prev_prices = [None] + prices[:-1]
    for i, (p, pp) in enumerate(zip(prices, prev_prices)):
        clr = (0, 200/255, 83/255, 0.12) if (pp is None or p >= pp) else (1, 23/255, 68/255, 0.12)
        ax.bar(i, p, color=clr, width=0.75, zorder=1)

    # Main line
    col_rgb = _hex_to_rgb01(color)
    ax.plot(x_idx, prices, color=col_rgb, linewidth=2.2, zorder=3)

    # Dots colored by direction
    for i, (p, pp) in enumerate(zip(prices, prev_prices)):
        dot_c = G if (pp is None or p >= pp) else R
        ax.scatter(i, p, color=dot_c, s=28, zorder=4, linewidths=0.8,
                   edgecolors=CARD)

    # Value labels
    for i, p in enumerate(prices):
        ax.text(i, p * 1.002, f"{p:,.0f}", ha="center", va="bottom",
                fontsize=7, color=SEC)

    # Forecast dotted
    if len(forecast):
        fc_labels = list(forecast["Label"])
        fc_prices = list(forecast["Price"])
        fc_x = [len(labels) - 1] + list(range(len(labels), len(labels) + len(fc_labels)))
        fc_y = [prices[-1]] + fc_prices
        ax.plot(fc_x, fc_y, color=GOLD_C, linewidth=1.5,
                linestyle="--", zorder=3)
        for i, (xi, p) in enumerate(zip(fc_x[1:], fc_prices)):
            ax.scatter(xi, p, color=GOLD_C, marker="D", s=30, zorder=4)
            ax.text(xi, p * 1.002, f"{p:,.0f}", ha="center", va="bottom",
                    fontsize=7, color=GOLD_C)
        all_labels = labels + fc_labels
    else:
        all_labels = labels

    ax.set_xticks(range(len(all_labels)))
    ax.set_xticklabels(all_labels, rotation=35, ha="right",
                       fontsize=8, color=SEC)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(colors=SEC, labelsize=8)
    ax.spines[:].set_color(GRD)
    ax.yaxis.label.set_color(SEC)
    ax.grid(axis="y", color=GRD, linewidth=0.5, zorder=0)
    ax.grid(axis="x", visible=False)

    # Legend
    from matplotlib.lines import Line2D
    legend_els = [
        Line2D([0], [0], color=col_rgb, linewidth=2, label=f"{metal} Monthly Avg"),
        Line2D([0], [0], color=GOLD_C,  linewidth=1.5, linestyle="--", label="Forecast ▸"),
    ]
    ax.legend(handles=legend_els, facecolor=CARD, edgecolor=GRD,
              labelcolor=SEC, fontsize=8, loc="upper left")

    ax.text(0.0, -0.18, "Source: LME", transform=ax.transAxes,
            fontsize=7.5, color="#484F58")

    fig.tight_layout(pad=0.8)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=CARD, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _mpl_quarterly_png(quarterly: pd.DataFrame, forecast_q: pd.DataFrame,
                       metal: str, color: str) -> bytes:
    """Dark-themed quarterly bar chart rendered by matplotlib — returns PNG bytes."""
    BG   = "#0D1117"
    CARD = "#161B22"
    GRD  = "#21262D"
    SEC  = "#8B949E"
    PRI  = "#E6EDF3"
    G    = "#00C853"
    R    = "#FF1744"
    GOLD_C = "#FFD600"

    col_rgb = _hex_to_rgb01(color)
    fig, ax = plt.subplots(figsize=(9.2, 4.2), facecolor=CARD)
    ax.set_facecolor(CARD)

    labels = list(quarterly["Label"])
    prices = list(quarterly["Price"])
    prev   = [None] + prices[:-1]
    bar_colors = [G if (pp is None or p >= pp) else R
                  for p, pp in zip(prices, prev)]
    x_idx = list(range(len(labels)))

    bars = ax.bar(x_idx, prices, color=bar_colors, alpha=0.85,
                  width=0.55, zorder=2)

    for i, p in enumerate(prices):
        ax.text(i, p * 1.003, f"{p:,.0f}", ha="center", va="bottom",
                fontsize=8.5, color=PRI)

    if len(forecast_q):
        fc_labels = list(forecast_q["Label"])
        fc_prices = list(forecast_q["Price"])
        fc_x = list(range(len(labels), len(labels) + len(fc_labels)))
        ax.bar(fc_x, fc_prices, color=GOLD_C, alpha=0.5,
               width=0.55, zorder=2, hatch="//", edgecolor=GOLD_C)
        for xi, p in zip(fc_x, fc_prices):
            ax.text(xi, p * 1.003, f"{p:,.0f}", ha="center", va="bottom",
                    fontsize=8.5, color=GOLD_C)
        all_labels = labels + fc_labels
        all_x      = x_idx + fc_x
    else:
        all_labels = labels
        all_x      = x_idx

    ax.set_xticks(all_x)
    ax.set_xticklabels(all_labels, rotation=20, ha="right",
                       fontsize=8, color=SEC)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(colors=SEC, labelsize=8)
    ax.spines[:].set_color(GRD)
    ax.grid(axis="y", color=GRD, linewidth=0.5, zorder=0)
    ax.grid(axis="x", visible=False)

    ax.text(0.0, -0.18, "Source: LME", transform=ax.transAxes,
            fontsize=7.5, color="#484F58")

    fig.tight_layout(pad=0.8)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, facecolor=CARD, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# PPTX BUILDER
# ─────────────────────────────────────────────
def build_pptx(cu_m, cu_q, cu_fm, cu_fq,
               al_m, al_q, al_fm, al_fq, sel_months) -> bytes:
    def _filter(df):
        r = df[df["Label"].isin(sel_months)]
        return r if not r.empty else df.tail(10)
    def _filter_q(df, mf):
        if mf.empty: return df.tail(4)
        s, e = mf["Date"].min(), mf["Date"].max()
        r    = df[(df["Date"] >= s) & (df["Date"] <= e)]
        return r if not r.empty else df.tail(4)

    cu_mf  = _filter(cu_m);  al_mf  = _filter(al_m)
    cu_qf  = _filter_q(cu_q, cu_mf); al_qf  = _filter_q(al_q, al_mf)
    cu_fcm = cu_fm[cu_fm["Label"].isin(sel_months)] if len(cu_fm) else cu_fm
    al_fcm = al_fm[al_fm["Label"].isin(sel_months)] if len(al_fm) else al_fm

    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(5.625)

    # Slide 1 — Cu Monthly
    _build_slide(prs,
        "Copper (Cu) — Monthly Average Price",
        f"USD/MT  ·  {cu_mf['Label'].iloc[0]} → {cu_mf['Label'].iloc[-1]}  ·  LME Cash Settlement",
        _mpl_monthly_png(cu_mf, pd.DataFrame(columns=["Date","Price","Label"]), "Copper (Cu)", COPPER_CLR),
        _kpi_dict(cu_mf), COPPER_CLR)

    # Slide 2 — Al Monthly
    _build_slide(prs,
        "Aluminium (Al) — Monthly Average Price",
        f"USD/MT  ·  {al_mf['Label'].iloc[0]} → {al_mf['Label'].iloc[-1]}  ·  LME Cash Settlement",
        _mpl_monthly_png(al_mf, pd.DataFrame(columns=["Date","Price","Label"]), "Aluminium (Al)", ALUM_CLR),
        _kpi_dict(al_mf), ALUM_CLR)

    # Slide 3 — Cu Quarterly
    kp3 = _kpi_dict(cu_mf)
    if len(cu_qf):
        kp3["l1"] = f"LATEST QUARTER ({cu_qf['Label'].iloc[-1]})"
        kp3["v1"] = f"${cu_qf['Price'].iloc[-1]:,.0f} /MT"
    _build_slide(prs,
        "Copper (Cu) — Quarterly Average Price",
        f"USD/MT  ·  {cu_qf['Label'].iloc[0] if len(cu_qf) else ''} → {cu_qf['Label'].iloc[-1] if len(cu_qf) else ''}",
        _mpl_quarterly_png(cu_qf, pd.DataFrame(columns=["Date","Price","Label"]), "Copper (Cu)", COPPER_CLR),
        kp3, COPPER_CLR)

    # Slide 4 — Al Quarterly
    kp4 = _kpi_dict(al_mf)
    if len(al_qf):
        kp4["l1"] = f"LATEST QUARTER ({al_qf['Label'].iloc[-1]})"
        kp4["v1"] = f"${al_qf['Price'].iloc[-1]:,.0f} /MT"
    _build_slide(prs,
        "Aluminium (Al) — Quarterly Average Price",
        f"USD/MT  ·  {al_qf['Label'].iloc[0] if len(al_qf) else ''} → {al_qf['Label'].iloc[-1] if len(al_qf) else ''}",
        _mpl_quarterly_png(al_qf, pd.DataFrame(columns=["Date","Price","Label"]), "Aluminium (Al)", ALUM_CLR),
        kp4, ALUM_CLR)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    # ── load ─────────────────────────────────
    with st.spinner(""):
        cu_raw = load_sheet(SHEET_COPPER)
        al_raw = load_sheet(SHEET_ALUMINIUM)

    if cu_raw.empty or al_raw.empty:
        st.error("No data loaded — check credentials / sheet names."); return

    cu_m = monthly_avg(cu_raw);     al_m = monthly_avg(al_raw)
    cu_q = quarterly_avg(cu_raw);   al_q = quarterly_avg(al_raw)
    cu_fm= linear_forecast(cu_m,3); al_fm= linear_forecast(al_m,3)
    cu_fq= linear_forecast(cu_q,2); al_fq= linear_forecast(al_q,2)

    # ── ticker bar ────────────────────────────
    render_ticker_bar(cu_m, al_m)

    # ── sidebar ───────────────────────────────
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:12px 0 14px">
          <div style="font-size:26px;font-weight:900;color:#FFFFFF;
                      letter-spacing:3px;font-family:'Arial Black',Arial,sans-serif;
                      line-height:1">
            LME<span style="color:#B87333">.</span>
          </div>
          <div style="font-size:9px;color:{TEXT_SEC};letter-spacing:2px;
                      text-transform:uppercase;margin-top:4px">
            London Metal Exchange
          </div>
          <div style="font-size:9px;color:{TEXT_MUT};letter-spacing:1.5px;
                      text-transform:uppercase;margin-top:2px">
            Metals Price Terminal
          </div>
        </div>
        <div style="height:2px;background:linear-gradient(90deg,#B87333,transparent);
                    margin-bottom:16px;border-radius:1px"></div>
        """, unsafe_allow_html=True)

        st.markdown(f'<div style="font-size:10px;color:{TEXT_SEC};'
                    f'letter-spacing:1.5px;text-transform:uppercase;'
                    f'margin-bottom:8px">Export Report</div>', unsafe_allow_html=True)

        all_months = sorted(
            set(cu_m["Label"].tolist() + al_m["Label"].tolist()),
            key=lambda x: datetime.strptime(x, "%b-%y"),
        )
        default_sel = all_months[-10:] if len(all_months) >= 10 else all_months

        sel_months = st.multiselect(
            "Select Months", options=all_months, default=default_sel,
            help="Choose months for the PPTX report (slides auto-filter)"
        )

        st.markdown(f"""
        <div style="font-size:10px;color:{TEXT_MUT};margin:8px 0 12px;
                    line-height:1.6">
          🖼 Slide 1 · Copper Monthly<br>
          🖼 Slide 2 · Aluminium Monthly<br>
          🖼 Slide 3 · Copper Quarterly<br>
          🖼 Slide 4 · Aluminium Quarterly
        </div>
        """, unsafe_allow_html=True)

        gen = st.button("⬇  Generate PPTX", type="primary")
        if gen:
            if not sel_months:
                st.warning("Select at least one month.")
            else:
                with st.spinner("Rendering slides…"):
                    try:
                        data = build_pptx(
                            cu_m, cu_q, cu_fm, cu_fq,
                            al_m, al_q, al_fm, al_fq, sel_months)
                        st.download_button(
                            "⬇  Download PPTX",
                            data=data,
                            file_name=f"LME_{datetime.now():%Y%m%d}.pptx",
                            mime="application/vnd.openxmlformats-officedocument"
                                 ".presentationml.presentation",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")

        st.markdown(f"""
        <div style="border-top:1px solid {BORDER};margin-top:16px;
                    padding-top:12px;font-size:10px;color:{TEXT_MUT}">
          <b style="color:{TEXT_SEC}">DATA COVERAGE</b><br><br>
          Cu: {cu_raw["Date"].min():%d %b %Y}<br>
          &nbsp;&nbsp;&nbsp;&nbsp;→ {cu_raw["Date"].max():%d %b %Y}<br>
          &nbsp;&nbsp;&nbsp;&nbsp;{len(cu_raw):,} daily records<br><br>
          Al: {al_raw["Date"].min():%d %b %Y}<br>
          &nbsp;&nbsp;&nbsp;&nbsp;→ {al_raw["Date"].max():%d %b %Y}<br>
          &nbsp;&nbsp;&nbsp;&nbsp;{len(al_raw):,} daily records
        </div>
        """, unsafe_allow_html=True)

    # ── tabs ──────────────────────────────────
    tab_cu, tab_al, tab_comp, tab_data = st.tabs([
        "  🟤  COPPER (Cu)  ",
        "  ⚙️  ALUMINIUM (Al)  ",
        "  📊  COMPARISON  ",
        "  📋  DATA TABLE  ",
    ])

    # ─ COPPER ────────────────────────────────
    with tab_cu:
        c1, c2 = st.columns([3.2, 1])
        with c1:
            st.plotly_chart(chart_live_daily(cu_raw, "Copper (Cu)", COPPER_CLR),
                            use_container_width=True, key="cu_live")
        with c2:
            render_kpis(cu_m)

        st.markdown(f'<div class="section-header">QUARTERLY & ROLLING AVERAGES</div>',
                    unsafe_allow_html=True)
        r1, r2 = st.columns(2)
        with r1:
            st.plotly_chart(quarterly_chart(cu_q, cu_fq, "Copper (Cu)", COPPER_CLR),
                            use_container_width=True, key="cu_q")
        with r2:
            st.plotly_chart(rolling_chart(cu_raw, "Copper (Cu)", COPPER_CLR),
                            use_container_width=True, key="cu_r")

        st.markdown(f'<div class="section-header" style="margin-top:20px">'
                    f'MONTHLY PRICE TABLE</div>', unsafe_allow_html=True)
        render_monthly_table(cu_m, COPPER_CLR)

    # ─ ALUMINIUM ─────────────────────────────
    with tab_al:
        a1, a2 = st.columns([3.2, 1])
        with a1:
            st.plotly_chart(chart_live_daily(al_raw, "Aluminium (Al)", ALUM_CLR),
                            use_container_width=True, key="al_live")
        with a2:
            render_kpis(al_m)

        st.markdown(f'<div class="section-header">QUARTERLY & ROLLING AVERAGES</div>',
                    unsafe_allow_html=True)
        ar1, ar2 = st.columns(2)
        with ar1:
            st.plotly_chart(quarterly_chart(al_q, al_fq, "Aluminium (Al)", ALUM_CLR),
                            use_container_width=True, key="al_q")
        with ar2:
            st.plotly_chart(rolling_chart(al_raw, "Aluminium (Al)", ALUM_CLR),
                            use_container_width=True, key="al_r")

        st.markdown(f'<div class="section-header" style="margin-top:20px">'
                    f'MONTHLY PRICE TABLE</div>', unsafe_allow_html=True)
        render_monthly_table(al_m, ALUM_CLR)

    # ─ COMPARISON ────────────────────────────
    with tab_comp:
        st.markdown(f'<div class="section-header">INDEXED PERFORMANCE (Base=100)</div>',
                    unsafe_allow_html=True)
        cu_idx = cu_m.copy(); cu_idx["Indexed"] = cu_idx["Price"] / cu_idx["Price"].iloc[0] * 100
        al_idx = al_m.copy(); al_idx["Indexed"] = al_idx["Price"] / al_idx["Price"].iloc[0] * 100

        fig_comp = go.Figure()
        fig_comp.add_trace(go.Scatter(
            x=cu_idx["Date"], y=cu_idx["Indexed"],
            line=dict(color=COPPER_CLR, width=2),
            mode="lines", name="Copper",
        ))
        fig_comp.add_trace(go.Scatter(
            x=al_idx["Date"], y=al_idx["Indexed"],
            line=dict(color=ALUM_CLR, width=2),
            mode="lines", name="Aluminium",
        ))
        fig_comp.add_hline(y=100, line_dash="dot", line_color=TEXT_MUT, line_width=1)
        fig_comp.update_layout(
            **_CHART_LAYOUT,
            title=dict(text="<b>Indexed Performance Comparison</b>  "
                            f"<span style='color:{TEXT_SEC};font-size:12px'>"
                            f"Base = 100 at start of period</span>",
                       font=dict(size=15)),
            height=380,
        )
        st.plotly_chart(fig_comp, use_container_width=True, key="comp_idx")

        st.markdown(f'<div class="section-header" style="margin-top:8px">'
                    f'SPREAD  (COPPER minus ALUMINIUM  ·  USD/MT)</div>',
                    unsafe_allow_html=True)
        merged = pd.merge(
            cu_m[["Date","Label","Price"]].rename(columns={"Price":"Cu"}),
            al_m[["Date","Label","Price"]].rename(columns={"Price":"Al"}),
            on=["Date","Label"], how="inner"
        )
        merged["Spread"] = merged["Cu"] - merged["Al"]
        fig_sp = go.Figure()
        fig_sp.add_trace(go.Bar(
            x=merged["Label"], y=merged["Spread"],
            marker_color=[GREEN if v >= 0 else RED for v in merged["Spread"]],
            name="Spread",
        ))
        fig_sp.update_layout(**_CHART_LAYOUT,
                             title=dict(text="<b>Cu–Al Spread</b>",
                                        font=dict(size=15)), height=300)
        st.plotly_chart(fig_sp, use_container_width=True, key="spread")

    # ─ DATA TABLE ────────────────────────────
    with tab_data:
        d1, d2 = st.columns(2)
        with d1:
            st.markdown(f'<div class="section-header">COPPER — DAILY PRICES</div>',
                        unsafe_allow_html=True)
            disp = cu_raw[["Date","Price"]].copy()
            disp["Date"]  = disp["Date"].dt.strftime("%d %b %Y")
            disp["Price"] = disp["Price"].apply(lambda x: f"${x:,.2f}")
            disp.columns  = ["Date", "Cash Price (USD/MT)"]
            st.dataframe(disp, height=400, use_container_width=True,
                         hide_index=True)
        with d2:
            st.markdown(f'<div class="section-header">ALUMINIUM — DAILY PRICES</div>',
                        unsafe_allow_html=True)
            disp2 = al_raw[["Date","Price"]].copy()
            disp2["Date"]  = disp2["Date"].dt.strftime("%d %b %Y")
            disp2["Price"] = disp2["Price"].apply(lambda x: f"${x:,.2f}")
            disp2.columns  = ["Date", "Cash Price (USD/MT)"]
            st.dataframe(disp2, height=400, use_container_width=True,
                         hide_index=True)

        st.markdown(f'<div class="section-header" style="margin-top:20px">'
                    f'QUARTERLY SUMMARY</div>', unsafe_allow_html=True)
        q_sum = pd.merge(
            cu_q[["Label","Price"]].rename(columns={"Price":"Cu Avg (USD/MT)"}),
            al_q[["Label","Price"]].rename(columns={"Price":"Al Avg (USD/MT)"}),
            on="Label", how="outer"
        ).sort_values("Label")
        q_sum["Cu Avg (USD/MT)"] = q_sum["Cu Avg (USD/MT)"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "—")
        q_sum["Al Avg (USD/MT)"] = q_sum["Al Avg (USD/MT)"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "—")
        st.dataframe(q_sum.set_index("Label"), use_container_width=True)

        st.markdown(f'<div class="section-header" style="margin-top:16px">'
                    f'FORECAST — NEXT 3 MONTHS (LINEAR)</div>', unsafe_allow_html=True)
        fc = pd.merge(
            cu_fm[["Label","Price"]].rename(columns={"Price":"Cu Forecast (USD/MT)"}),
            al_fm[["Label","Price"]].rename(columns={"Price":"Al Forecast (USD/MT)"}),
            on="Label", how="outer",
        )
        fc["Cu Forecast (USD/MT)"] = fc["Cu Forecast (USD/MT)"].apply(lambda x: f"${x:,.2f}")
        fc["Al Forecast (USD/MT)"] = fc["Al Forecast (USD/MT)"].apply(lambda x: f"${x:,.2f}")
        st.dataframe(fc.set_index("Label"), use_container_width=True)


if __name__ == "__main__":
    main()
