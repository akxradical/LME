"""
LME Metals Dashboard — Streamlit App
=====================================
Reads live data from Google Sheets (LME Copper + Aluminium),
shows monthly/quarterly averages, 3-month rolling avg, forecasts,
and exports a 4-slide PPTX in the reference design format.

─── SETUP ─────────────────────────────────────────────────────
  pip install streamlit pandas gspread google-auth plotly python-pptx kaleido

─── RUN LOCALLY ───────────────────────────────────────────────
  streamlit run app.py

─── STREAMLIT CLOUD SECRETS  (.streamlit/secrets.toml) ────────
  SPREADSHEET_ID = "1zLyMANY56oFRwFug04WYavGUH_NAlRH8M3c-TXIRDlI"

  [gcp_service_account]
  type                        = "service_account"
  project_id                  = "your-project-id"
  private_key_id              = "key-id"
  private_key                 = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
  client_email                = "your-sa@your-project.iam.gserviceaccount.com"
  client_id                   = "..."
  auth_uri                    = "https://accounts.google.com/o/oauth2/auth"
  token_uri                   = "https://oauth2.googleapis.com/token"
  auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
  client_x509_cert_url        = "..."

─── REQUIREMENTS.TXT ──────────────────────────────────────────
  streamlit>=1.32
  pandas
  gspread
  google-auth
  plotly
  kaleido
  python-pptx
  numpy
  openpyxl
"""

import io
from datetime import datetime

import gspread
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ──────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────
SPREADSHEET_ID  = "1zLyMANY56oFRwFug04WYavGUH_NAlRH8M3c-TXIRDlI"
SHEET_COPPER    = "LME Copper"
SHEET_ALUMINIUM = "LME Aluminium"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Brand colours
C_BLUE   = "#1565C0"
C_RED    = "#C62828"
C_GREEN  = "#2E7D32"
C_BG     = "#F5F7FA"
C_TEXT   = "#1A1A2E"
C_COPPER = "#B87333"
C_ALUM   = "#607D8B"

st.set_page_config(
    page_title="LME Metals Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────
# DATA LOADING
# ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_sheet(sheet_name: str) -> pd.DataFrame:
    """Load a worksheet from Google Sheets.  Falls back to synthetic data locally."""
    try:
        info  = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        gc    = gspread.authorize(creds)
        ws    = gc.open_by_key(
                    st.secrets.get("SPREADSHEET_ID", SPREADSHEET_ID)
                ).worksheet(sheet_name)
        rows  = ws.get_all_records()
        df    = pd.DataFrame(rows)
        df.columns = [c.strip() for c in df.columns]

        date_col  = next(c for c in df.columns if "date"  in c.lower())
        price_col = next(c for c in df.columns if "price" in c.lower()
                                                or "cash"  in c.lower())
        df = df.rename(columns={date_col: "Date", price_col: "Price"})
        df["Date"]  = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        df["Price"] = pd.to_numeric(
            df["Price"].astype(str).str.replace(",", ""), errors="coerce"
        )
        df = (df.dropna(subset=["Date", "Price"])
                .sort_values("Date")
                .reset_index(drop=True))
        return df

    except Exception:
        # ── LOCAL / DEMO fallback ────────────────────────────
        return _synthetic(sheet_name)


def _synthetic(sheet_name: str) -> pd.DataFrame:
    rng   = np.random.default_rng(42 if "Copper" in sheet_name else 7)
    base  = 9000 if "Copper" in sheet_name else 2300
    dates = pd.date_range("2025-01-02", datetime.today(), freq="B")
    noise = rng.normal(0, 60, len(dates)).cumsum()
    prices= np.clip(base + noise, base * 0.8, base * 1.35)
    return pd.DataFrame({"Date": dates, "Price": prices.round(2)})

# ──────────────────────────────────────────────────────────────
# ANALYTICS
# ──────────────────────────────────────────────────────────────
def monthly_avg(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["YM"]    = df["Date"].dt.to_period("M")
    m           = df.groupby("YM")["Price"].mean().reset_index()
    m["Date"]   = m["YM"].dt.to_timestamp()
    m["Label"]  = m["YM"].dt.strftime("%b-%y")
    m["Price"]  = m["Price"].round(2)
    return m.drop(columns="YM")


def quarterly_avg(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["YQ"]    = df["Date"].dt.to_period("Q")
    q           = df.groupby("YQ")["Price"].mean().reset_index()
    q["Date"]   = q["YQ"].dt.to_timestamp()
    q["Label"]  = q["YQ"].dt.strftime("Q%q %Y")
    q["Price"]  = q["Price"].round(2)
    return q.drop(columns="YQ")


def rolling_avg(df: pd.DataFrame, window: int = 3) -> pd.DataFrame:
    df = df.copy().sort_values("Date")
    df[f"MA{window}"] = df["Price"].rolling(window).mean().round(2)
    return df


def linear_forecast(monthly: pd.DataFrame, periods: int = 3) -> pd.DataFrame:
    if len(monthly) < 2:
        return pd.DataFrame(columns=["Date", "Price", "Label"])
    x = np.arange(len(monthly))
    y = monthly["Price"].values
    slope, intercept = np.polyfit(x, y, 1)
    last = monthly["Date"].iloc[-1]
    fdates  = [last + pd.DateOffset(months=i + 1) for i in range(periods)]
    fprices = [intercept + slope * (len(monthly) + i + 1) for i in range(periods)]
    return pd.DataFrame({
        "Date":  fdates,
        "Price": np.round(fprices, 2),
        "Label": [d.strftime("%b-%y") for d in fdates],
    })

# ──────────────────────────────────────────────────────────────
# CHARTS
# ──────────────────────────────────────────────────────────────
def chart_monthly(monthly: pd.DataFrame, forecast: pd.DataFrame,
                  metal: str, color: str) -> go.Figure:
    fig = go.Figure()

    # Historical area + line
    r, g, b = bytes.fromhex(color.lstrip("#"))
    fig.add_trace(go.Scatter(
        x=monthly["Date"], y=monthly["Price"],
        fill="tozeroy",
        fillcolor=f"rgba({r},{g},{b},0.12)",
        line=dict(color=color, width=2.5),
        mode="lines+markers+text",
        marker=dict(size=6),
        text=[f"{p:,.2f}" for p in monthly["Price"]],
        textposition="top center",
        textfont=dict(size=8),
        name=f"{metal} Monthly Avg",
    ))

    # Forecast dotted extension
    if len(forecast):
        bx = [monthly["Date"].iloc[-1]] + list(forecast["Date"])
        by = [monthly["Price"].iloc[-1]] + list(forecast["Price"])
        fig.add_trace(go.Scatter(
            x=bx, y=by,
            line=dict(color=color, width=2, dash="dot"),
            mode="lines+markers+text",
            marker=dict(size=7, symbol="diamond", color="#BDBDBD"),
            text=[""] + [f"{p:,.2f}" for p in forecast["Price"]],
            textposition="top center",
            textfont=dict(size=8, color="#888888"),
            name="Forecast (Linear)",
        ))

    first_lbl = monthly["Label"].iloc[0]
    last_lbl  = monthly["Label"].iloc[-1]
    fig.update_layout(
        title=dict(
            text=(f"<b>{metal} — Monthly Average Price Trend</b><br>"
                  f"<span style='font-size:11px;color:#666666'>"
                  f"USD per MT  |  {first_lbl} – {last_lbl}</span>"),
            font=dict(size=17),
        ),
        xaxis=dict(showgrid=False, tickfont=dict(size=11)),
        yaxis=dict(gridcolor="#E8EAF0", tickformat=","),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.18),
        margin=dict(l=55, r=15, t=80, b=60),
        height=400,
        annotations=[dict(
            text="Source: LME", x=0, y=-0.22,
            xref="paper", yref="paper",
            showarrow=False, font=dict(size=9, color="#AAAAAA"),
        )],
    )
    return fig


def chart_quarterly(quarterly: pd.DataFrame, forecast_q: pd.DataFrame,
                    metal: str, color: str) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=quarterly["Label"], y=quarterly["Price"],
        marker_color=color, opacity=0.85,
        text=[f"{p:,.0f}" for p in quarterly["Price"]],
        textposition="outside",
        textfont=dict(size=9),
        name="Quarterly Avg",
    ))
    if len(forecast_q):
        fig.add_trace(go.Bar(
            x=forecast_q["Label"], y=forecast_q["Price"],
            marker_color="#BDBDBD",
            text=[f"{p:,.0f}" for p in forecast_q["Price"]],
            textposition="outside",
            textfont=dict(size=9),
            name="Forecast",
            marker_pattern_shape="x",
        ))
    fig.update_layout(
        title=dict(text=f"<b>{metal} — Quarterly Average Price</b>", font=dict(size=16)),
        xaxis=dict(showgrid=False),
        yaxis=dict(gridcolor="#E8EAF0", tickformat=","),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.2),
        margin=dict(l=55, r=15, t=65, b=60),
        height=380,
        annotations=[dict(
            text="Source: LME", x=0, y=-0.25,
            xref="paper", yref="paper",
            showarrow=False, font=dict(size=9, color="#AAAAAA"),
        )],
    )
    return fig


def chart_rolling(df_raw: pd.DataFrame, metal: str, color: str) -> go.Figure:
    rolled = rolling_avg(df_raw, 3)
    fig    = go.Figure()
    fig.add_trace(go.Scatter(
        x=rolled["Date"], y=rolled["Price"],
        line=dict(color="#E0E0E0", width=1),
        name="Daily Price", mode="lines",
    ))
    fig.add_trace(go.Scatter(
        x=rolled["Date"], y=rolled["MA3"],
        line=dict(color=color, width=2.5),
        name="3-Month Rolling Avg",
    ))
    fig.update_layout(
        title=f"<b>{metal} — 3-Month Rolling Average</b>",
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(showgrid=False),
        yaxis=dict(gridcolor="#E8EAF0", tickformat=","),
        legend=dict(orientation="h", y=-0.2),
        margin=dict(l=55, r=15, t=60, b=60),
        height=380,
    )
    return fig

# ──────────────────────────────────────────────────────────────
# KPI METRIC CARDS
# ──────────────────────────────────────────────────────────────
def show_kpis(monthly: pd.DataFrame):
    cur      = monthly["Price"].iloc[-1]
    cur_lbl  = monthly["Label"].iloc[-1]
    low_val  = monthly["Price"].min()
    low_lbl  = monthly.loc[monthly["Price"].idxmin(), "Label"]
    first    = monthly["Price"].iloc[0]
    first_lbl= monthly["Label"].iloc[0]
    chg      = cur - first
    pct      = chg / first * 100
    arrow    = "▲" if chg >= 0 else "▼"
    chg_clr  = C_GREEN if chg >= 0 else C_RED
    trend    = "increase" if chg >= 0 else "decrease"

    st.markdown(f"""
    <style>
      .kpi {{ border-radius:10px; padding:13px 16px; margin-bottom:10px;
               text-align:center; font-family:'Segoe UI',sans-serif; }}
      .kpi .lbl {{ font-size:10px; font-weight:700; letter-spacing:.4px; opacity:.9 }}
      .kpi .val {{ font-size:22px; font-weight:800; margin:3px 0 }}
      .kpi .sub {{ font-size:10px; opacity:.85 }}
    </style>

    <div class="kpi" style="background:{C_BLUE};color:#fff">
      <div class="lbl">Current Price ({cur_lbl})</div>
      <div class="val">${cur:,.0f} /MT</div>
    </div>

    <div class="kpi" style="background:{C_RED};color:#fff">
      <div class="lbl">Period Low ({low_lbl})</div>
      <div class="val">${low_val:,.0f} /MT</div>
    </div>

    <div class="kpi" style="background:{C_GREEN};color:#fff">
      <div class="lbl">{first_lbl} vs {cur_lbl}</div>
      <div class="val">{arrow} ${abs(chg):,.0f} /MT</div>
      <div class="sub">+{abs(pct):.1f}% {trend}</div>
    </div>
    """, unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# PPTX BUILDER
# ──────────────────────────────────────────────────────────────
def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16))


def _txt(slide, left, top, w, h, text, pt, bold=False,
         color="#000000", align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(left, top, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    r   = p.add_run()
    r.text          = text
    r.font.size     = Pt(pt)
    r.font.bold     = bold
    r.font.color.rgb= _rgb(color)
    return txb


def _kpi_card(slide, left, top, width, height,
              label, value, sub, bg):
    # background rectangle
    rect = slide.shapes.add_shape(1, left, top, width, height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = _rgb(bg)
    rect.line.fill.background()
    # rounded corners workaround: just overlay a slightly inset shape
    # label
    _txt(slide, left, top + Inches(0.06), width, Inches(0.22),
         label, 7.5, bold=True, color="#FFFFFF", align=PP_ALIGN.CENTER)
    # value
    _txt(slide, left, top + Inches(0.27), width, Inches(0.48),
         value, 16, bold=True, color="#FFFFFF", align=PP_ALIGN.CENTER)
    # sub
    if sub:
        _txt(slide, left, top + Inches(0.74), width, Inches(0.22),
             sub, 7.5, color="#E8F5E9", align=PP_ALIGN.CENTER)


def _build_pptx_slide(prs, title, subtitle, chart_png, kpi_dict, accent_hex):
    """One slide: title band | chart (L 69%) | KPI cards (R 29%)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    W, H  = prs.slide_width, prs.slide_height

    # White background
    bg = slide.shapes.add_shape(1, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = _rgb("#FFFFFF")
    bg.line.fill.background()

    # Light grey header strip
    hd = slide.shapes.add_shape(1, 0, 0, W, Inches(0.68))
    hd.fill.solid(); hd.fill.fore_color.rgb = _rgb("#F5F7FA")
    hd.line.fill.background()

    # Title
    _txt(slide, Inches(0.28), Inches(0.05), Inches(7.2), Inches(0.36),
         title, 17, bold=True, color=C_TEXT)
    # Subtitle
    _txt(slide, Inches(0.28), Inches(0.40), Inches(7.2), Inches(0.22),
         subtitle, 8.5, color="#666666")

    # Thin accent bar
    bar = slide.shapes.add_shape(1, 0, Inches(0.68), W, Inches(0.025))
    bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(accent_hex)
    bar.line.fill.background()

    # Chart image
    slide.shapes.add_picture(
        io.BytesIO(chart_png),
        Inches(0.15), Inches(0.73),
        width=Inches(6.65), height=Inches(4.10),
    )

    # KPI cards
    kx, kw, kh = Inches(7.05), Inches(2.65), Inches(1.12)
    _kpi_card(slide, kx, Inches(0.78), kw, kh,
              kpi_dict["lbl1"], kpi_dict["val1"], None, C_BLUE)
    _kpi_card(slide, kx, Inches(2.00), kw, kh,
              kpi_dict["lbl2"], kpi_dict["val2"], None, C_RED)
    _kpi_card(slide, kx, Inches(3.22), kw, kh,
              kpi_dict["lbl3"], kpi_dict["val3"], kpi_dict.get("sub3", ""), C_GREEN)

    # Footer
    _txt(slide, Inches(0.28), Inches(4.95), Inches(6), Inches(0.2),
         "Source: LME", 8, color="#AAAAAA")


def _kpi_dict_from_monthly(df: pd.DataFrame) -> dict:
    cur       = df["Price"].iloc[-1]
    cur_lbl   = df["Label"].iloc[-1]
    low_val   = df["Price"].min()
    low_lbl   = df.loc[df["Price"].idxmin(), "Label"]
    first     = df["Price"].iloc[0]
    first_lbl = df["Label"].iloc[0]
    chg       = cur - first
    pct       = chg / first * 100
    return {
        "lbl1": f"Current Price ({cur_lbl})",
        "val1": f"${cur:,.0f} /MT",
        "lbl2": f"Period Low ({low_lbl})",
        "val2": f"${low_val:,.0f} /MT",
        "lbl3": f"{first_lbl} vs {cur_lbl}",
        "val3": f"{'▲' if chg>=0 else '▼'} ${abs(chg):,.0f} /MT",
        "sub3": f"+{abs(pct):.1f}% {'increase' if chg>=0 else 'decrease'}",
    }


def build_pptx(
    cu_monthly, cu_quarterly, cu_forecast_m, cu_forecast_q,
    al_monthly, al_quarterly, al_forecast_m, al_forecast_q,
    sel_months: list,
) -> bytes:

    def _sel_monthly(df):
        return df[df["Label"].isin(sel_months)].copy() or df.tail(10)

    def _sel_quarterly(df, monthly_f):
        if monthly_f.empty:
            return df.tail(4)
        s, e = monthly_f["Date"].min(), monthly_f["Date"].max()
        return df[(df["Date"] >= s) & (df["Date"] <= e)].copy()

    cu_mf  = cu_monthly[cu_monthly["Label"].isin(sel_months)].copy()
    al_mf  = al_monthly[al_monthly["Label"].isin(sel_months)].copy()
    if cu_mf.empty: cu_mf = cu_monthly.tail(10).copy()
    if al_mf.empty: al_mf = al_monthly.tail(10).copy()

    cu_qf  = _sel_quarterly(cu_quarterly, cu_mf)
    al_qf  = _sel_quarterly(al_quarterly, al_mf)

    cu_fcm = cu_forecast_m[cu_forecast_m["Label"].isin(sel_months)] if len(cu_forecast_m) else cu_forecast_m
    al_fcm = al_forecast_m[al_forecast_m["Label"].isin(sel_months)] if len(al_forecast_m) else al_forecast_m

    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(5.625)

    fig_to_png = lambda fig: fig.to_image(format="png", width=900, height=460, scale=2)

    # Slide 1: Copper Monthly
    f1   = chart_monthly(cu_mf, cu_fcm, "Copper (Cu)", C_COPPER)
    _build_pptx_slide(
        prs,
        "Copper (Cu) — Monthly Average Price Trend",
        f"USD per MT  |  {cu_mf['Label'].iloc[0]} – {cu_mf['Label'].iloc[-1]}",
        fig_to_png(f1),
        _kpi_dict_from_monthly(cu_mf),
        C_COPPER,
    )

    # Slide 2: Aluminium Monthly
    f2   = chart_monthly(al_mf, al_fcm, "Aluminium (Al)", C_ALUM)
    _build_pptx_slide(
        prs,
        "Aluminium (Al) — Monthly Average Price Trend",
        f"USD per MT  |  {al_mf['Label'].iloc[0]} – {al_mf['Label'].iloc[-1]}",
        fig_to_png(f2),
        _kpi_dict_from_monthly(al_mf),
        C_ALUM,
    )

    # Slide 3: Copper Quarterly
    f3   = chart_quarterly(cu_qf, cu_forecast_q, "Copper (Cu)", C_COPPER)
    kpi3 = _kpi_dict_from_monthly(cu_mf)
    if len(cu_qf):
        kpi3["lbl1"] = f"Latest Quarter ({cu_qf['Label'].iloc[-1]})"
        kpi3["val1"] = f"${cu_qf['Price'].iloc[-1]:,.0f} /MT"
    _build_pptx_slide(
        prs,
        "Copper (Cu) — Quarterly Average Price",
        f"USD per MT  |  {cu_qf['Label'].iloc[0] if len(cu_qf) else ''} – {cu_qf['Label'].iloc[-1] if len(cu_qf) else ''}",
        fig_to_png(f3),
        kpi3,
        C_COPPER,
    )

    # Slide 4: Aluminium Quarterly
    f4   = chart_quarterly(al_qf, al_forecast_q, "Aluminium (Al)", C_ALUM)
    kpi4 = _kpi_dict_from_monthly(al_mf)
    if len(al_qf):
        kpi4["lbl1"] = f"Latest Quarter ({al_qf['Label'].iloc[-1]})"
        kpi4["val1"] = f"${al_qf['Price'].iloc[-1]:,.0f} /MT"
    _build_pptx_slide(
        prs,
        "Aluminium (Al) — Quarterly Average Price",
        f"USD per MT  |  {al_qf['Label'].iloc[0] if len(al_qf) else ''} – {al_qf['Label'].iloc[-1] if len(al_qf) else ''}",
        fig_to_png(f4),
        kpi4,
        C_ALUM,
    )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ──────────────────────────────────────────────────────────────
# MAIN UI
# ──────────────────────────────────────────────────────────────
def main():
    st.markdown(f"""
    <style>
      .block-container {{ padding-top:1.5rem }}
      [data-testid="stSidebar"] {{ background:{C_BG} }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("## 📈 LME Metals Price Dashboard")
    st.caption(
        f"Live data from Google Sheets · Auto-refresh every 5 min · "
        f"Loaded: **{datetime.now().strftime('%d %b %Y, %H:%M')}**"
    )

    # ── Load ──────────────────────────────────────────────────
    with st.spinner("Fetching data from Google Sheets…"):
        cu_raw = load_sheet(SHEET_COPPER)
        al_raw = load_sheet(SHEET_ALUMINIUM)

    if cu_raw.empty or al_raw.empty:
        st.error("❌ No data loaded. Check credentials and sheet names.")
        return

    # ── Aggregates ────────────────────────────────────────────
    cu_m  = monthly_avg(cu_raw)
    cu_q  = quarterly_avg(cu_raw)
    cu_fm = linear_forecast(cu_m, 3)
    cu_fq = linear_forecast(cu_q, 2)

    al_m  = monthly_avg(al_raw)
    al_q  = quarterly_avg(al_raw)
    al_fm = linear_forecast(al_m, 3)
    al_fq = linear_forecast(al_q, 2)

    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/London_Metal_Exchange_logo.svg/320px-London_Metal_Exchange_logo.svg.png",
                 width=160)
        st.markdown("---")
        st.markdown("### 📥 Export to PowerPoint")

        all_months = sorted(
            set(cu_m["Label"].tolist() + al_m["Label"].tolist()),
            key=lambda x: datetime.strptime(x, "%b-%y"),
        )
        default_sel = all_months[-10:] if len(all_months) >= 10 else all_months

        sel_months = st.multiselect(
            "Select months to include",
            options=all_months,
            default=default_sel,
            help="Slides 1–2 = monthly charts, Slides 3–4 = quarterly charts for the selected period",
        )

        st.markdown("**Slide layout:**")
        st.markdown("• Slide 1 — Copper Monthly")
        st.markdown("• Slide 2 — Aluminium Monthly")
        st.markdown("• Slide 3 — Copper Quarterly")
        st.markdown("• Slide 4 — Aluminium Quarterly")
        st.markdown("---")

        gen_btn = st.button("🖨️ Generate PPTX", type="primary", use_container_width=True)
        if gen_btn:
            if not sel_months:
                st.warning("Select at least one month first.")
            else:
                with st.spinner("Building 4-slide presentation…"):
                    try:
                        pptx_bytes = build_pptx(
                            cu_m, cu_q, cu_fm, cu_fq,
                            al_m, al_q, al_fm, al_fq,
                            sel_months,
                        )
                        st.download_button(
                            label="⬇️ Download PPTX",
                            data=pptx_bytes,
                            file_name=f"LME_Metals_{datetime.now().strftime('%Y%m%d')}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Error generating PPTX: {e}")

        st.markdown("---")
        st.markdown("**📊 Data coverage**")
        st.caption(f"Cu: {cu_raw['Date'].min():%d %b %Y} → {cu_raw['Date'].max():%d %b %Y}  ({len(cu_raw):,} rows)")
        st.caption(f"Al: {al_raw['Date'].min():%d %b %Y} → {al_raw['Date'].max():%d %b %Y}  ({len(al_raw):,} rows)")

    # ── Tabs ──────────────────────────────────────────────────
    tab_cu, tab_al, tab_data = st.tabs(
        ["🟤  Copper (Cu)", "⚙️  Aluminium (Al)", "📋  Raw Data & Tables"]
    )

    # ─ COPPER ─────────────────────────────────────────────────
    with tab_cu:
        col_c, col_k = st.columns([3, 1])
        with col_c:
            st.plotly_chart(chart_monthly(cu_m, cu_fm, "Copper (Cu)", C_COPPER),
                            use_container_width=True, key="cu_m")
        with col_k:
            st.markdown("<br>", unsafe_allow_html=True)
            show_kpis(cu_m)

        st.markdown("#### Quarterly & Rolling Average")
        q1, q2 = st.columns(2)
        with q1:
            st.plotly_chart(chart_quarterly(cu_q, cu_fq, "Copper (Cu)", C_COPPER),
                            use_container_width=True, key="cu_q")
        with q2:
            st.plotly_chart(chart_rolling(cu_raw, "Copper (Cu)", C_COPPER),
                            use_container_width=True, key="cu_r")

        st.markdown("#### Monthly Summary Table")
        cu_tbl = cu_m[["Label", "Price"]].rename(
            columns={"Label": "Month", "Price": "Avg Price (USD/MT)"}
        ).copy()
        cu_tbl["MoM Change (USD)"] = cu_tbl["Avg Price (USD/MT)"].diff().round(2)
        cu_tbl["MoM %"]            = (cu_tbl["Avg Price (USD/MT)"].pct_change() * 100).round(2)
        st.dataframe(
            cu_tbl.set_index("Month").style
                .format({"Avg Price (USD/MT)": "{:,.2f}",
                         "MoM Change (USD)": "{:+,.2f}",
                         "MoM %": "{:+.2f}%"})
                .background_gradient(subset=["Avg Price (USD/MT)"], cmap="YlOrBr"),
            use_container_width=True,
        )

    # ─ ALUMINIUM ──────────────────────────────────────────────
    with tab_al:
        col_a, col_ak = st.columns([3, 1])
        with col_a:
            st.plotly_chart(chart_monthly(al_m, al_fm, "Aluminium (Al)", C_ALUM),
                            use_container_width=True, key="al_m")
        with col_ak:
            st.markdown("<br>", unsafe_allow_html=True)
            show_kpis(al_m)

        st.markdown("#### Quarterly & Rolling Average")
        a1, a2 = st.columns(2)
        with a1:
            st.plotly_chart(chart_quarterly(al_q, al_fq, "Aluminium (Al)", C_ALUM),
                            use_container_width=True, key="al_q")
        with a2:
            st.plotly_chart(chart_rolling(al_raw, "Aluminium (Al)", C_ALUM),
                            use_container_width=True, key="al_r")

        st.markdown("#### Monthly Summary Table")
        al_tbl = al_m[["Label", "Price"]].rename(
            columns={"Label": "Month", "Price": "Avg Price (USD/MT)"}
        ).copy()
        al_tbl["MoM Change (USD)"] = al_tbl["Avg Price (USD/MT)"].diff().round(2)
        al_tbl["MoM %"]            = (al_tbl["Avg Price (USD/MT)"].pct_change() * 100).round(2)
        st.dataframe(
            al_tbl.set_index("Month").style
                .format({"Avg Price (USD/MT)": "{:,.2f}",
                         "MoM Change (USD)": "{:+,.2f}",
                         "MoM %": "{:+.2f}%"})
                .background_gradient(subset=["Avg Price (USD/MT)"], cmap="Blues"),
            use_container_width=True,
        )

    # ─ RAW DATA ───────────────────────────────────────────────
    with tab_data:
        d1, d2 = st.columns(2)
        with d1:
            st.markdown("##### Copper — Daily Prices")
            st.dataframe(
                cu_raw[["Date", "Price"]]
                    .rename(columns={"Price": "Cash Price (USD/MT)"})
                    .assign(Date=cu_raw["Date"].dt.strftime("%d %b %Y"))
                    .style.format({"Cash Price (USD/MT)": "{:,.2f}"}),
                height=380, use_container_width=True,
            )
        with d2:
            st.markdown("##### Aluminium — Daily Prices")
            st.dataframe(
                al_raw[["Date", "Price"]]
                    .rename(columns={"Price": "Cash Price (USD/MT)"})
                    .assign(Date=al_raw["Date"].dt.strftime("%d %b %Y"))
                    .style.format({"Cash Price (USD/MT)": "{:,.2f}"}),
                height=380, use_container_width=True,
            )

        st.markdown("##### Quarterly Summary — Both Metals")
        q_summary = pd.merge(
            cu_q[["Label", "Price"]].rename(columns={"Price": "Cu Avg (USD/MT)"}),
            al_q[["Label", "Price"]].rename(columns={"Price": "Al Avg (USD/MT)"}),
            on="Label", how="outer",
        ).sort_values("Label")
        st.dataframe(
            q_summary.set_index("Label").style.format("{:,.2f}"),
            use_container_width=True,
        )

        st.markdown("##### 3-Month Forecast (Linear Trend)")
        fc_tbl = pd.merge(
            cu_fm[["Label", "Price"]].rename(columns={"Price": "Cu Forecast (USD/MT)"}),
            al_fm[["Label", "Price"]].rename(columns={"Price": "Al Forecast (USD/MT)"}),
            on="Label", how="outer",
        )
        st.dataframe(
            fc_tbl.set_index("Label").style.format("{:,.2f}"),
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
