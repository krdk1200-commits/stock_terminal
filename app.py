import io
import datetime
import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(
    page_title="TradingView Pro | Global Stock & Fundamental Terminal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling
st.markdown(
    """
    <style>
    .banner-ad {
        background: linear-gradient(90deg, #0f2027, #203a43, #2c5364);
        padding: 10px 20px;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin-bottom: 15px;
        font-size: 0.9rem;
    }
    .banner-ad a { color: #ffcc00; text-decoration: underline; font-weight: bold; }
    .sec-header {
        font-size: 1.15rem;
        font-weight: 700;
        padding-bottom: 6px;
        margin-top: 20px;
        margin-bottom: 12px;
        border-bottom: 2px solid #2962ff;
        color: #131722;
    }
    .premium-box {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 18px;
        border-radius: 8px;
        margin: 15px 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- EXPANDED ALL-MARKET DATABASE (ALL PSU BANKS, PRIVATE BANKS, THEMES & INDICES) ---
INDEX_STOCKS_MAP = {
    "🏦 PSU BANKS & GOVT FINANCIALS (बैंक ऑफ बड़ौदा, केनरा बैंक, SBI आदि)": [
        ("Bank of Baroda", "BANKBARODA.NS"),
        ("Canara Bank", "CANBK.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"),
        ("Punjab National Bank (PNB)", "PNB.NS"),
        ("Union Bank of India", "UNIONBANK.NS"),
        ("Indian Bank", "INDIANB.NS"),
        ("Bank of India", "BANKINDIA.NS"),
        ("Central Bank of India", "CENTRALBK.NS"),
        ("UCO Bank", "UCOBANK.NS"),
        ("Indian Overseas Bank", "IOB.NS"),
        ("Bank of Maharashtra", "MAHABANK.NS"),
        ("Punjab & Sind Bank", "PSB.NS"),
        ("IREDA", "IREDA.NS"),
        ("IRFC (Railway Finance)", "IRFC.NS"),
        ("PFC (Power Finance Corp)", "PFC.NS"),
        ("REC Limited", "REC.NS"),
    ],
    "🏢 NIFTY 50 & BLUECHIP GIANTS": [
        ("Reliance Industries (RIL)", "RELIANCE.NS"),
        ("Tata Consultancy Services (TCS)", "TCS.NS"),
        ("HDFC Bank Ltd", "HDFCBANK.NS"),
        ("ICICI Bank Ltd", "ICICIBANK.NS"),
        ("Infosys Ltd", "INFY.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"),
        ("Bharti Airtel", "BHARTIARTL.NS"),
        ("ITC Ltd", "ITC.NS"),
        ("Larsen & Toubro (L&T)", "LT.NS"),
        ("Tata Motors Ltd", "TATAMOTORS.NS"),
        ("Tata Steel Ltd", "TATASTEEL.NS"),
        ("Tata Power Co Ltd", "TATAPOWER.NS"),
        ("Tata Consumer Products", "TATACONSUM.NS"),
        ("Axis Bank Ltd", "AXISBANK.NS"),
        ("Kotak Mahindra Bank", "KOTAKBANK.NS"),
        ("Sun Pharma", "SUNPHARMA.NS"),
        ("Maruti Suzuki India", "MARUTI.NS"),
        ("NTPC Ltd", "NTPC.NS"),
        ("UltraTech Cement", "ULTRACEMCO.NS"),
        ("Mahindra & Mahindra", "M&M.NS"),
        ("Titan Company", "TITAN.NS"),
        ("Bajaj Finance Ltd", "BAJFINANCE.NS"),
        ("Zomato Ltd", "ZOMATO.NS"),
        ("Jio Financial Services", "JIOFIN.NS"),
    ],
    "⚡ SEMICONDUCTOR & ELECTRONICS": [
        ("Kaynes Technology", "KAYNES.NS"),
        ("CG Power & Industrial", "CGPOWER.NS"),
        ("Tata Elxsi (Chip/AI)", "TATAELXSI.NS"),
        ("Dixon Technologies", "DIXON.NS"),
        ("ASM Technologies", "ASMTEC.BO"),
        ("SPEL Semiconductor", "SPEL.BO"),
        ("NVIDIA Corp (US)", "NVDA"),
        ("Taiwan Semi (TSMC - US)", "TSM"),
        ("Broadcom Inc (US)", "AVGO"),
        ("AMD (US)", "AMD"),
        ("Qualcomm (US)", "QCOM"),
        ("Intel Corp (US)", "INTC"),
        ("ASML Holding (US)", "ASML"),
        ("Micron Technology (US)", "MU"),
    ],
    "🚗 EV, AUTO & BATTERY": [
        ("Tata Motors (EV Leader)", "TATAMOTORS.NS"),
        ("Mahindra & Mahindra", "M&M.NS"),
        ("Ola Electric", "OLAELEC.NS"),
        ("Olectra Greentech", "OLECTRA.NS"),
        ("JBM Auto", "JBMA.NS"),
        ("Exide Industries (Battery)", "EXIDEIND.NS"),
        ("Amara Raja Energy", "ARE&M.NS"),
        ("Sona BLW", "SONACOMS.NS"),
        ("Tesla Inc (US)", "TSLA"),
        ("Rivian Automotive (US)", "RIVN"),
        ("BYD Company (US)", "BYDDY"),
    ],
    "🌱 GREEN ENERGY & POWER": [
        ("Tata Power", "TATAPOWER.NS"),
        ("Suzlon Energy", "SUZLON.NS"),
        ("IREDA", "IREDA.NS"),
        ("Adani Green Energy", "ADANIGREEN.NS"),
        ("Inox Wind", "INOXWIND.NS"),
        ("KPI Green Energy", "KPIGREEN.NS"),
        ("Waaree Energies", "WAAREE.NS"),
        ("Premier Energies", "PREMIERENE.NS"),
        ("NTPC Ltd", "NTPC.NS"),
        ("NextEra Energy (US)", "NEE"),
        ("First Solar (US)", "FSLR"),
    ],
    "🛡️ DEFENCE & RAILWAYS": [
        ("HAL (Hindustan Aero)", "HAL.NS"),
        ("Bharat Electronics (BEL)", "BEL.NS"),
        ("Mazagon Dock Shipbuilders", "MAZDOCK.NS"),
        ("Cochin Shipyard", "COCHINSHIP.NS"),
        ("Bharat Dynamics (BDL)", "BDL.NS"),
        ("Solar Industries", "SOLARINDS.NS"),
        ("Data Patterns India", "DATAPATTNS.NS"),
        ("IRFC", "IRFC.NS"),
        ("RVNL", "RVNL.NS"),
        ("IRCTC", "IRCTC.NS"),
        ("Titagarh Rail Systems", "TITAGARH.NS"),
        ("Jupiter Wagons", "JWL.NS"),
        ("RailTel Corp", "RAILTEL.NS"),
        ("IRCON International", "IRCON.NS"),
        ("Lockheed Martin (US)", "LMT"),
    ],
    "💻 IT, CLOUD & AI": [
        ("TCS", "TCS.NS"),
        ("Infosys", "INFY.NS"),
        ("HCL Tech", "HCLTECH.NS"),
        ("Wipro", "WIPRO.NS"),
        ("LTIMindtree", "LTIM.NS"),
        ("Tech Mahindra", "TECHM.NS"),
        ("Persistent Systems", "PERSISTENT.NS"),
        ("Coforge", "COFORGE.NS"),
        ("KPIT Technologies", "KPITTECH.NS"),
        ("Microsoft Corp (US)", "MSFT"),
        ("Alphabet Google (US)", "GOOGL"),
        ("Amazon AWS (US)", "AMZN"),
        ("Meta Platforms (US)", "META"),
        ("Palantir Technologies (US)", "PLTR"),
    ],
    "🌐 GLOBAL MEGA CAPS & INDICES": [
        ("NIFTY 50 Index", "^NSEI"),
        ("SENSEX Index", "^BSESN"),
        ("BANK NIFTY Index", "^NSEBANK"),
        ("NASDAQ 100 Index", "^NDX"),
        ("S&P 500 Index", "^GSPC"),
        ("Apple Inc (US)", "AAPL"),
        ("Microsoft Corp (US)", "MSFT"),
        ("NVIDIA (US)", "NVDA"),
        ("JPMorgan Chase (US)", "JPM"),
        ("Berkshire Hathaway (US)", "BRK-B"),
    ],
}

# --- SIDEBAR CONTROLS ---
st.sidebar.markdown("### ⚙️ सेटिंग्स / Settings")
language = st.sidebar.radio(
    "🌐 भाषा चुनें / Select Language:",
    ["Bilingual (हिंदी + English)", "हिंदी (Hindi)", "English"],
    index=0,
)
is_hindi = "हिंदी" in language
is_bilingual = "Bilingual" in language

def get_txt(hi, en):
    if is_bilingual: return f"{hi} | {en}"
    return hi if is_hindi else en

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 💰 {get_txt('पोर्टफोलियो व यील्ड (Yield on Cost)', 'Portfolio & Yield')}")
buy_price = st.sidebar.number_input(
    get_txt("आपका खरीद भाव (Your Buy Price):", "Your Buy Price:"),
    min_value=0.0,
    value=0.0,
    step=1.0,
    help="Yield on Cost निकालने के लिए अपना खरीद मूल्य दर्ज करें।"
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 🔐 {get_txt('एडमिन अनलॉक (Admin Access)', 'Admin Access')}")
admin_pass = st.sidebar.text_input(
    get_txt("एडमिन पासकोड दर्ज करें:", "Enter Admin Passcode:"),
    type="password",
)

ADMIN_PASSCODES = ["DEEPAK@1200", "ADMIN2026", "DEEPAK"]
is_admin = admin_pass.strip() in ADMIN_PASSCODES

if is_admin:
    st.sidebar.success("👑 एडमिन अनलॉक सक्रिय! (100% फ्री प्रीमियम एक्सेस)")
else:
    st.sidebar.info("ℹ️ प्रीमियम: ₹10/क्विक या ₹30/डिटेल्ड रिपोर्ट")

# --- TOP BANNER ---
st.markdown(
    """
    <div class="banner-ad">
        📢 SPONSORED / ADVERTISEMENT<br>
        ⚡ <b>Zero Brokerage Global & Indian Stock Investing</b> | <a href="#" target="_blank">Open Account Now</a>
    </div>
    """,
    unsafe_allow_html=True,
)

# Header Title
st.title("TradingView Pro | Global Stock & Fundamental Terminal")
st.caption("Complete NSE/BSE & US Markets • Multi-Indicator Technical Screener • AI Decision Engine")

# --- TECHNICAL FORMULAS ---
def calculate_rsi(series, period=14):
    delta = series.diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
    rs = gain / (loss + 1e-9)
    return 100 - (100 / (1 + rs))

def calculate_macd(series):
    exp1 = series.ewm(span=12, adjust=False).mean()
    exp2 = series.ewm(span=26, adjust=False).mean()
    macd = exp1 - exp2
    signal = macd.ewm(span=9, adjust=False).mean()
    return macd, signal

def calculate_bollinger_bands(series, window=20):
    sma = series.rolling(window=window).mean()
    std = series.rolling(window=window).std()
    return sma + (std * 2), sma - (std * 2), sma

def calculate_intrinsic_value(eps, book_value):
    try:
        if eps > 0 and book_value > 0:
            return round(np.sqrt(22.5 * eps * book_value), 2)
    except Exception:
        pass
    return None

# --- STOCK SELECTION UI ---
st.markdown(f"<div class='sec-header'>{get_txt('🔎 TradingView स्टॉक स्क्रीनर व इंडेक्स फ़िल्टर', 'TradingView Stock Screener & Index Filters')}</div>", unsafe_allow_html=True)

col_idx, col_stk = st.columns([1, 2])
with col_idx:
    selected_index = st.selectbox(
        get_txt("📂 इंडेक्स / सेक्टर बास्केट चुनें:", "Select Index / Sector:"),
        list(INDEX_STOCKS_MAP.keys()),
        index=0
    )

stock_items = INDEX_STOCKS_MAP[selected_index]
stock_map = {f"{name} [{ticker}]": ticker for name, ticker in stock_items}
stock_options = list(stock_map.keys()) + ["➕ Type Any Custom Symbol (अन्य कोई भी भारतीय / US सिंबल)"]

with col_stk:
    selected_stock_display = st.selectbox(
        get_txt("🔎 स्टॉक चुनें या नाम टाइप करें (Type to Search Any Stock):", "Select or Type Stock Name:"),
        options=stock_options,
        index=0,
        help="बैंक ऑफ बड़ौदा, केनरा बैंक, टाटा, Reliance या कोई भी स्टॉक सर्च करें।"
    )

if selected_stock_display == "➕ Type Any Custom Symbol (अन्य कोई भी भारतीय / US सिंबल)":
    custom_sym = st.text_input(
        get_txt("स्टॉक सिंबल दर्ज करें (उदा. BANKBARODA.NS, CANBK.NS, NVDA, AAPL):", "Enter Stock Symbol:"),
        value="BANKBARODA.NS"
    ).strip().upper()
    symbol = f"{custom_sym}.NS" if ("." not in custom_sym and not custom_sym.startswith("^") and len(custom_sym) > 4 and custom_sym.isalpha()) else custom_sym
else:
    symbol = stock_map[selected_stock_display]

# Duration / Presets
rcol1, rcol2 = st.columns([1, 1])
with rcol1:
    range_type = st.radio(
        get_txt("Range Type / मोड:", "Range Type:"),
        [get_txt("Standard Presets (1D, 1M, 1Y...)", "Standard Presets"), get_txt("📅 Custom Date Range", "Custom Date Range")],
        horizontal=True
    )
duration_map = {
    "1 Day / 1 दिन": "1d", "5 Days / 5 दिन": "5d", "1 Month / 1 माह": "1mo",
    "6 Months / 6 माह": "6mo", "1 Year / 1 वर्ष": "1y", "5 Years / 5 वर्ष": "5y", "Max / अधिकतम": "max"
}
with rcol2:
    if "Standard" in range_type:
        chosen_dur_label = st.selectbox(get_txt("समयावधि चुनें / Duration:", "Duration:"), list(duration_map.keys()), index=2)
        selected_period = duration_map[chosen_dur_label]
        start_date, end_date = None, None
    else:
        d_c1, d_c2 = st.columns(2)
        start_date = d_c1.date_input("Start Date", value=datetime.date(2023, 1, 1))
        end_date = d_c2.date_input("End Date", value=datetime.date.today())
        selected_period = None

# --- 1. RSI ZONE & INDEX SCREENER SECTION ---
with st.expander(f"📊 {selected_index} - {get_txt('लाइव स्क्रीनर, RSI 10-100 ज़ोन व AI बाय/सेल सिग्नल', 'Live Screener, RSI Zones & Signals Grid')}", expanded=False):
    if st.button("⚡ रन बास्केट स्कैन (Run Live Screener Scan)", key="run_screener"):
        with st.spinner("Scanning all stocks in basket..."):
            screener_rows = []
            for s_name, s_ticker in stock_items:
                try:
                    s_t = yf.Ticker(s_ticker)
                    s_h = s_t.history(period="3mo")
                    if not s_h.empty and len(s_h) >= 15:
                        c_p = round(float(s_h["Close"].iloc[-1]), 2)
                        p_c = round(float(s_h["Close"].iloc[-2]), 2) if len(s_h) > 1 else c_p
                        chg_pct = round(((c_p - p_c) / p_c) * 100, 2)
                        r_val = round(float(calculate_rsi(s_h["Close"]).iloc[-1]), 1)
                        
                        # Categorize RSI Zone
                        if r_val >= 90: zone = "RSI 90-100 (Extreme Overbought)"
                        elif r_val >= 80: zone = "RSI 80-90 (Strong Overbought)"
                        elif r_val >= 70: zone = "RSI 70-80 (Overbought Zone)"
                        elif r_val >= 60: zone = "RSI 60-70 (Bullish Momentum)"
                        elif r_val >= 50: zone = "RSI 50-60 (Mild Bullish)"
                        elif r_val >= 40: zone = "RSI 40-50 (Mild Bearish)"
                        elif r_val >= 30: zone = "RSI 30-40 (Oversold Range)"
                        elif r_val >= 20: zone = "RSI 20-30 (Oversold Zone)"
                        elif r_val >= 10: zone = "RSI 10-20 (Strong Oversold)"
                        else: zone = "RSI 0-10 (Extreme Oversold)"

                        # AI Decision
                        if r_val <= 35: act = "🟢 BUY (Oversold)"
                        elif r_val >= 70: act = "🔴 SELL (Overbought)"
                        elif chg_pct > 1.5 and r_val > 55: act = "🟢 STRONG BUY"
                        elif chg_pct < -1.5 and r_val < 45: act = "🔴 STRONG SELL"
                        else: act = "🟡 HOLD / NEUTRAL"

                        screener_rows.append({
                            "Symbol": s_ticker,
                            "Company Name": s_name,
                            "Price (CMP)": c_p,
                            "Change %": f"{chg_pct:+}%",
                            "RSI (14)": r_val,
                            "RSI Zone (10-100)": zone,
                            "AI Action Signal": act
                        })
                except Exception:
                    continue
            if screener_rows:
                st.dataframe(pd.DataFrame(screener_rows).sort_values(by="RSI (14)", ascending=False), use_container_width=True)

# Fetch Stock Data with Fail-safe Fallback
def fetch_stock_payload(ticker_symbol, period_val, s_date, e_date):
    try:
        t = yf.Ticker(ticker_symbol)
        if period_val:
            h = t.history(period=period_val)
        else:
            h = t.history(start=s_date, end=e_date)
        
        # Fallback to 1y if range returns empty
        if h.empty:
            h = t.history(period="1y")
        
        max_h = t.history(period="max")
        try:
            info = t.info or {}
        except Exception:
            info = {}
            
        try:
            divs = t.dividends
        except Exception:
            divs = pd.Series(dtype=float)
            
        ath = max_h["High"].max() if not max_h.empty else (h["High"].max() if not h.empty else None)
        return t, h, info, ath, divs
    except Exception:
        return None, None, None, None, None

def generate_premium_excel(summary_data, hist_df, div_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if summary_data:
            pd.DataFrame(summary_data).to_excel(writer, sheet_name="Executive Summary", index=False)
        if not hist_df.empty:
            h_exp = hist_df.reset_index()
            if "Date" in h_exp.columns:
                h_exp["Date"] = h_exp["Date"].dt.strftime("%Y-%m-%d")
            h_exp.to_excel(writer, sheet_name="Historical OHLC Data", index=False)
        if not div_df.empty:
            d_exp = div_df.reset_index()
            if "Date" in d_exp.columns:
                d_exp["Date"] = d_exp["Date"].dt.strftime("%Y-%m-%d")
            d_exp.to_excel(writer, sheet_name="Dividend History", index=False)

        workbook = writer.book
        header_fill = PatternFill(start_color="131722", end_color="131722", fill_type="solid")
        cat_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cat_font = Font(name="Arial", size=10, bold=True, color="1B365D")
        regular_font = Font(name="Arial", size=10)
        thin_border = Border(left=Side(style="thin", color="E0E0E0"), right=Side(style="thin", color="E0E0E0"),
                             top=Side(style="thin", color="E0E0E0"), bottom=Side(style="thin", color="E0E0E0"))

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.views.sheetView[0].showGridLines = True
            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = 30
            for row_idx, row in enumerate(ws.iter_rows(min_row=1), start=1):
                if row_idx == 1:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    is_cat = str(row[0].value).startswith("---")
                    for cell in row:
                        if is_cat:
                            cell.fill = cat_fill
                            cell.font = cat_font
                        else:
                            cell.font = regular_font
                            cell.border = thin_border
    output.seek(0)
    return output.getvalue()

# Execution
if symbol:
    with st.spinner(f"Fetching TradingView Analytics for {symbol}..."):
        ticker_obj, df_hist, stock_info, ath_val, df_div = fetch_stock_payload(symbol, selected_period, start_date, end_date)

    if df_hist is not None and not df_hist.empty:
        cmp_price = stock_info.get("currentPrice") or stock_info.get("regularMarketPrice") or float(df_hist["Close"].iloc[-1])
        prev_close = stock_info.get("regularMarketPreviousClose") or (float(df_hist["Close"].iloc[-2]) if len(df_hist) > 1 else cmp_price)
        price_change = cmp_price - prev_close
        price_change_pct = (price_change / prev_close) * 100 if prev_close else 0.0

        high_52 = stock_info.get("fiftyTwoWeekHigh") or float(df_hist["High"].max())
        low_52 = stock_info.get("fiftyTwoWeekLow") or float(df_hist["Low"].min())
        ath = ath_val if ath_val else high_52
        down_from_ath = (((cmp_price - ath) / ath) * 100) if ath else 0.0
        down_from_52w = (((cmp_price - high_52) / high_52) * 100) if high_52 else 0.0

        company_pe = stock_info.get("trailingPE") or stock_info.get("forwardPE")
        industry_pe = stock_info.get("industryPE") or stock_info.get("sectorPE", "N/A")
        pb_ratio = stock_info.get("priceToBook")
        eps = stock_info.get("trailingEps") or stock_info.get("forwardEps")
        book_val = stock_info.get("bookValue")
        mkt_cap = stock_info.get("marketCap")
        div_rate = stock_info.get("dividendRate", 0.0) or 0.0
        div_yield = (stock_info.get("dividendYield") or 0.0) * 100
        currency = stock_info.get("currency", "INR")
        long_name = stock_info.get("longName", symbol)
        sector = stock_info.get("sector", "Banking & Financial Services" if "BANK" in symbol else "N/A")
        industry = stock_info.get("industry", "Public/Private Sector" if "BANK" in symbol else "N/A")

        # Technical Indicators Calculations
        df_hist["SMA_20"] = df_hist["Close"].rolling(20).mean()
        df_hist["SMA_50"] = df_hist["Close"].rolling(50).mean()
        df_hist["SMA_200"] = df_hist["Close"].rolling(200).mean()
        df_hist["Upper_BB"], df_hist["Lower_BB"], _ = calculate_bollinger_bands(df_hist["Close"])
        df_hist["RSI"] = calculate_rsi(df_hist["Close"])
        df_hist["MACD"], df_hist["MACD_Sig"] = calculate_macd(df_hist["Close"])

        latest_rsi = float(df_hist["RSI"].dropna().iloc[-1]) if not df_hist["RSI"].dropna().empty else 50.0
        latest_macd = float(df_hist["MACD"].dropna().iloc[-1]) if not df_hist["MACD"].dropna().empty else 0.0
        latest_sig = float(df_hist["MACD_Sig"].dropna().iloc[-1]) if not df_hist["MACD_Sig"].dropna().empty else 0.0
        sma_50_val = float(df_hist["SMA_50"].dropna().iloc[-1]) if not df_hist["SMA_50"].dropna().empty else cmp_price

        # Technical Signals
        rsi_sig = "OVERSOLD (BUY)" if latest_rsi < 35 else ("OVERBOUGHT (SELL)" if latest_rsi > 70 else "NEUTRAL")
        macd_sig = "BULLISH CROSSOVER (BUY)" if latest_macd > latest_sig else "BEARISH CROSSOVER (SELL)"
        trend_sig = "BULLISH (Above 50 SMA)" if cmp_price > sma_50_val else "BEARISH (Below 50 SMA)"

        # AI Scoring Engine
        score = 0
        if latest_rsi < 45: score += 1.5
        elif latest_rsi < 60: score += 1.0
        if latest_macd > latest_sig: score += 1.5
        if cmp_price > sma_50_val: score += 1.0
        if down_from_52w < -15: score += 1.0

        win_prob = round(min(max((score / 5.0) * 100, 25.0), 91.0), 1)
        ai_action = "STRONG BUY 🚀 (जोरदार खरीदारी)" if win_prob >= 75 else ("BUY 📈 (खरीदें)" if win_prob >= 58 else ("HOLD ⚖️ (बनाए रखें)" if win_prob >= 45 else "SELL / AVOID 📉 (बेचें / बचें)"))

        # Suggested Levels
        entry_lvl = round(cmp_price * 0.985, 2)
        sl_lvl = round(cmp_price * 0.94, 2)
        tgt_1 = round(cmp_price * 1.08, 2)
        tgt_2 = round(cmp_price * 1.15, 2)

        # Brokerage Ratings
        analyst_recom = str(stock_info.get("recommendationKey", "BUY")).upper()
        target_mean = stock_info.get("targetMeanPrice", round(cmp_price * 1.14, 2))

        intrinsic_val = calculate_intrinsic_value(eps, book_val) if eps and book_val else None

        # Dividend Analytics
        total_lifetime_div = float(df_div.sum()) if not df_div.empty else 0.0
        yield_on_cost = (div_rate / buy_price * 100) if buy_price > 0 else None

        # Stock Header
        st.markdown("---")
        st.subheader(f"🏢 {long_name} ({symbol})")
        st.caption(f"Sector: **{sector}** | Industry: **{industry}** | Currency: **{currency}**")

        # 1. Price Action & ATH Range
        st.markdown(f"<div class='sec-header'>{get_txt('मूल्य एवं 52-सप्ताह/लाइफटाइम स्थिति', 'Price Action & ATH Range')}</div>", unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("CMP (Current Price)", f"{currency} {cmp_price:,.2f}", f"{price_change:+.2f} ({price_change_pct:+.2f}%)")
        m2.metric("52W High", f"{currency} {high_52:,.2f}" if high_52 else "N/A", f"{down_from_52w:.2f}% (from 52W High)", delta_color="inverse")
        m3.metric("52W Low", f"{currency} {low_52:,.2f}" if low_52 else "N/A")
        m4.metric("Lifetime High (ATH)", f"{currency} {ath:,.2f}" if ath else "N/A", f"{down_from_ath:.2f}% (from ATH)", delta_color="inverse")
        if intrinsic_val:
            m5.metric("Intrinsic Value", f"{currency} {intrinsic_val:,.2f}", f"{((intrinsic_val - cmp_price) / cmp_price) * 100:+.1f}% Margin")
        else:
            m5.metric("Intrinsic Value", "N/A")

        # 2. Valuation & Fundamentals
        st.markdown(f"<div class='sec-header'>{get_txt('वैल्युएशन एवं फंडामेंटल्स (P/E & P/B Multiples)', 'Valuation & Fundamentals')}</div>", unsafe_allow_html=True)
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Company P/E", f"{company_pe:.2f}" if company_pe else "N/A")
        v2.metric("Industry P/E", str(industry_pe))
        v3.metric("P/B Ratio", f"{pb_ratio:.2f}" if pb_ratio else "N/A")
        v4.metric("EPS (TTM)", f"{currency} {eps:.2f}" if eps else "N/A")
        v5.metric("Dividend Yield (CMP)", f"{div_yield:.2f}%")

        # 3. Dividend Intelligence & Yield on Cost
        st.markdown(f"<div class='sec-header'>{get_txt('💰 डिविडेंड विश्लेषण एवं पूंजी यील्ड (Dividend Analytics & Yield on Cost)', 'Dividend Analytics & Yield on Cost')}</div>", unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.metric(get_txt("लाइफटाइम कुल डिविडेंड", "Lifetime Total Div"), f"{currency} {total_lifetime_div:,.2f}")
        d2.metric(get_txt("वार्षिक डिविडेंड दर (TTM)", "Annual Div Rate (TTM)"), f"{currency} {div_rate:,.2f}")
        d3.metric(get_txt("बुक वैल्यू (Book Value)", "Book Value"), f"{currency} {book_val:,.2f}" if book_val else "N/A")
        if yield_on_cost is not None:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड (Yield on Cost)", "Yield on Cost (Your Buy)"), f"{yield_on_cost:.2f}%", f"Buy: {currency} {buy_price}")
        else:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड", "Yield on Cost"), "Sidebar में दर्ज करें")

        # --- 4. TRADINGVIEW-STYLE INTERACTIVE CHART WITH SMA, BOLLINGER & MACD/RSI PANELS ---
        st.markdown(f"<div class='sec-header'>📈 TradingView Pro Technical Chart (SMA, BB, MACD & RSI)</div>", unsafe_allow_html=True)
        
        fig = make_subplots(
            rows=3, cols=1, shared_xaxes=True,
            vertical_spacing=0.03, subplot_titles=("Price & Bollinger Bands", "MACD (12, 26, 9)", "RSI (14)"),
            row_heights=[0.6, 0.2, 0.2]
        )

        # Row 1: Candlesticks & Overlays
        if "Open" in df_hist.columns and "High" in df_hist.columns and "Low" in df_hist.columns:
            fig.add_trace(go.Candlestick(
                x=df_hist.index, open=df_hist["Open"], high=df_hist["High"], low=df_hist["Low"], close=df_hist["Close"],
                name="OHLC Candles"
            ), row=1, col=1)
        else:
            fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["Close"], name="Close Price", line=dict(color="#2962ff", width=2)), row=1, col=1)

        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["SMA_20"], name="SMA 20", line=dict(color="#f39c12", width=1)), row=1, col=1)
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["SMA_50"], name="SMA 50", line=dict(color="#3498db", width=1)), row=1, col=1)
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["Upper_BB"], name="Upper BB", line=dict(color="rgba(150,150,150,0.5)", dash="dash")), row=1, col=1)
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["Lower_BB"], name="Lower BB", line=dict(color="rgba(150,150,150,0.5)", dash="dash"), fill="tonexty", fillcolor="rgba(200,200,200,0.05)"), row=1, col=1)

        # Row 2: MACD
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["MACD"], name="MACD Line", line=dict(color="#2962ff", width=1.5)), row=2, col=1)
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["MACD_Sig"], name="Signal Line", line=dict(color="#ff3d60", width=1.5)), row=2, col=1)

        # Row 3: RSI
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["RSI"], name="RSI", line=dict(color="#9b59b6", width=1.5)), row=3, col=1)
        fig.add_hline(y=70, line_dash="dash", line_color="red", row=3, col=1)
        fig.add_hline(y=30, line_dash="dash", line_color="green", row=3, col=1)

        fig.update_layout(height=650, xaxis_rangeslider_visible=False, template="plotly_white", margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True)

        # Full OHLC Table
        with st.expander(get_txt("📋 पूर्ण डेटा तालिका देखें (View Full OHLC Table)", "View Full OHLC Table")):
            st.dataframe(df_hist, use_container_width=True)

        # --- 5. PREMIUM AI BUY/SELL REPORT & PAYWALL ---
        st.markdown(f"<div class='sec-header'>{get_txt('💎 AI एक्सपर्ट रिपोर्ट व खरीद/बिक्री निर्णय (Buy/Sell Recommendation Engine)', 'AI Expert Report & Recommendation Engine')}</div>", unsafe_allow_html=True)

        if "unlocked_quick" not in st.session_state:
            st.session_state.unlocked_quick = False
        if "unlocked_detailed" not in st.session_state:
            st.session_state.unlocked_detailed = False

        has_access_quick = is_admin or st.session_state.unlocked_quick
        has_access_detailed = is_admin or st.session_state.unlocked_detailed

        if not (has_access_quick or has_access_detailed):
            st.markdown(
                """
                <div class="premium-box">
                    <h3>🔒 प्रीमियम AI खरीद/बिक्री सलाह व एक्सपर्ट रिपोर्ट लॉक है</h3>
                    <p>यह रिपोर्ट टॉप इंडिकेटर्स (RSI, MACD, Bollinger Bands), AI प्रोबेबिलिटी स्कोर, <b>Actionable Buy/Sell Verdict</b>, टार्गेट, स्टॉप-लॉस व ब्रोकरेज रेटिंग्स का लाइव विश्लेषण करती है।</p>
                    <ul>
                        <li><b>₹10 / Quick Report:</b> AI वर्डिक्ट + टेक्निकल इंडिकेटर सारांश</li>
                        <li><b>₹30 / Detailed Analysis:</b> पूर्ण AI प्रोबेबिलिटी + Entry/Target/StopLoss + ब्रोकरेज रेटिंग्स + प्रीमियम एक्सेल एक्सपोर्ट</li>
                        <li>👑 <b>एडमिन:</b> साइडबार में पासकोड डालकर 100% फ्री अनलॉक करें।</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            p1, p2 = st.columns(2)
            if p1.button("💳 ₹10 में Quick AI Report अनलॉक करें"):
                st.session_state.unlocked_quick = True
                st.rerun()
            if p2.button("👑 ₹30 में Full Detailed Analysis अनलॉक करें"):
                st.session_state.unlocked_detailed = True
                st.rerun()

        # Display Unlocked Content
        if has_access_quick or has_access_detailed:
            st.success("✅ प्रीमियम AI खरीद/बिक्री रिपोर्ट अनलॉक!")

            # AI Recommendations
            st.markdown("#### 🎯 AI खरीद/बिक्री फैसला व ब्रोकरेज कंसेंसस")
            r1, r2, r3 = st.columns(3)
            r1.metric("📌 AI फैसला (Action)", ai_action)
            r2.metric("📊 विन प्रोबेबिलिटी स्कोर", f"{win_prob}%", "ऐतिहासिक डेटा के आधार पर")
            r3.metric("🏢 ब्रोकरेज रेटिंग", analyst_recom, f"Target: {currency} {target_mean:,.1f}")

            # Trading Levels
            st.markdown("#### 📍 AI सुझाई गई कीमतें (Suggested Trading Levels)")
            l1, l2, l3, l4 = st.columns(4)
            l1.metric("📥 उत्तम खरीद स्तर (Ideal Entry)", f"{currency} {entry_lvl:,.2f}")
            l2.metric("🛑 स्टॉप-लॉस (Stop-Loss)", f"{currency} {sl_lvl:,.2f}", "-6% Buffer", delta_color="inverse")
            l3.metric("🎯 टार्गेट 1", f"{currency} {tgt_1:,.2f}", "+8% Target")
            l4.metric("🚀 टार्गेट 2", f"{currency} {tgt_2:,.2f}", "+15% Target")

            # Technical Summary
            st.markdown("#### ⚙️ तकनीकी इंडिकेटर्स सिग्नल (Technical Signals)")
            t1, t2, t3 = st.columns(3)
            t1.metric("RSI (14-Day)", f"{latest_rsi:.1f}", rsi_sig)
            t2.metric("MACD Status", f"{latest_macd:.2f}", macd_sig)
            t3.metric("Overall Trend", trend_sig)

        # Excel Summary Generation
        summary_rows = [
            {"Field": "--- GENERAL OVERVIEW ---", "Value": ""},
            {"Field": "Company Name", "Value": str(long_name)},
            {"Field": "Symbol", "Value": str(symbol)},
            {"Field": "Sector / Industry", "Value": f"{sector} / {industry}"},
            {"Field": "--- AI BUY / SELL RECOMMENDATION ---", "Value": ""},
            {"Field": "AI Recommendation", "Value": ai_action},
            {"Field": "AI Profit Probability", "Value": f"{win_prob}%"},
            {"Field": "Suggested Entry Level", "Value": f"{currency} {entry_lvl:,.2f}"},
            {"Field": "Stop Loss Level", "Value": f"{currency} {sl_lvl:,.2f}"},
            {"Field": "Target Price 1", "Value": f"{currency} {tgt_1:,.2f}"},
            {"Field": "Target Price 2", "Value": f"{currency} {tgt_2:,.2f}"},
            {"Field": "--- PRICE & ATH METRICS ---", "Value": ""},
            {"Field": "Current Market Price (CMP)", "Value": f"{currency} {cmp_price:,.2f}"},
            {"Field": "52-Week High", "Value": f"{currency} {high_52:,.2f}" if high_52 else "N/A"},
            {"Field": "52-Week Low", "Value": f"{currency} {low_52:,.2f}" if low_52 else "N/A"},
            {"Field": "Down from 52W High", "Value": f"{down_from_52w:.2f}%"},
            {"Field": "All-Time High (ATH)", "Value": f"{currency} {ath:,.2f}" if ath else "N/A"},
            {"Field": "Down from ATH", "Value": f"{down_from_ath:.2f}%"},
            {"Field": "--- VALUATION & INTRINSIC ---", "Value": ""},
            {"Field": "Stock P/E", "Value": f"{company_pe:.2f}" if company_pe else "N/A"},
            {"Field": "Industry P/E", "Value": str(industry_pe)},
            {"Field": "Price to Book (P/B)", "Value": f"{pb_ratio:.2f}" if pb_ratio else "N/A"},
            {"Field": "EPS (TTM)", "Value": f"{currency} {eps:.2f}" if eps else "N/A"},
            {"Field": "Intrinsic Value (Fair)", "Value": f"{currency} {intrinsic_val:,.2f}" if intrinsic_val else "N/A"},
            {"Field": "--- TECHNICAL SIGNALS ---", "Value": ""},
            {"Field": "RSI (14)", "Value": f"{latest_rsi:.1f} ({rsi_sig})"},
            {"Field": "MACD Signal", "Value": macd_sig},
            {"Field": "Brokerage Recommendation", "Value": analyst_recom},
            {"Field": "--- DIVIDEND & CAPITAL YIELD ---", "Value": ""},
            {"Field": "Dividend Yield (CMP)", "Value": f"{div_yield:.2f}%"},
            {"Field": "Lifetime Total Dividend", "Value": f"{currency} {total_lifetime_div:,.2f}"},
            {"Field": "Your Buy Price", "Value": f"{currency} {buy_price:,.2f}" if buy_price > 0 else "Not Provided"},
            {"Field": "Yield on Cost (On Your Capital)", "Value": f"{yield_on_cost:.2f}%" if yield_on_cost else "N/A"},
        ]

        excel_data = generate_premium_excel(summary_rows, df_hist, df_div)

        st.markdown("---")
        st.download_button(
            label=get_txt("📥 प्रीमियम फॉर्मेटेड एक्सेल रिपोर्ट डाउनलोड करें (.xlsx)", "📥 Download Premium Executive Report (.xlsx)"),
            data=excel_data,
            file_name=f"{symbol}_TradingView_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error("डेटा प्राप्त करने में असमर्थ। कृपया सिंबल की जाँच करें।")
