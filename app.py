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

# Custom Styling (TradingView inspired Dark/Light UI)
st.markdown(
    """
    <style>
    .tradingview-header {
        background-color: #131722;
        color: #d1d4dc;
        padding: 10px 20px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 15px;
    }
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

# --- COMPREHENSIVE INDEX & STOCK BASKETS ---
INDEX_STOCKS_MAP = {
    "NIFTY 50": [
        ("Reliance Industries", "RELIANCE.NS"), ("TCS", "TCS.NS"), ("HDFC Bank", "HDFCBANK.NS"),
        ("ICICI Bank", "ICICIBANK.NS"), ("Infosys", "INFY.NS"), ("State Bank of India", "SBIN.NS"),
        ("Bharti Airtel", "BHARTIARTL.NS"), ("ITC Ltd", "ITC.NS"), ("Larsen & Toubro", "LT.NS"),
        ("Tata Motors", "TATAMOTORS.NS"), ("Tata Steel", "TATASTEEL.NS"), ("Axis Bank", "AXISBANK.NS"),
        ("Kotak Bank", "KOTAKBANK.NS"), ("Sun Pharma", "SUNPHARMA.NS"), ("Maruti Suzuki", "MARUTI.NS"),
        ("NTPC Ltd", "NTPC.NS"), ("UltraTech Cement", "ULTRACEMCO.NS"), ("Mahindra & Mahindra", "M&M.NS"),
        ("Titan Company", "TITAN.NS"), ("Bajaj Finance", "BAJFINANCE.NS")
    ],
    "NIFTY BANK & PSU BANKS": [
        ("State Bank of India", "SBIN.NS"), ("Bank of Baroda", "BANKBARODA.NS"), ("Canara Bank", "CANBK.NS"),
        ("Punjab National Bank", "PNB.NS"), ("Union Bank of India", "UNIONBANK.NS"), ("Indian Bank", "INDIANB.NS"),
        ("HDFC Bank", "HDFCBANK.NS"), ("ICICI Bank", "ICICIBANK.NS"), ("Axis Bank", "AXISBANK.NS"),
        ("Kotak Mahindra Bank", "KOTAKBANK.NS"), ("Federal Bank", "FEDERALBNK.NS"), ("IDFC First Bank", "IDFCFIRSTB.NS")
    ],
    "NIFTY IT & TECH": [
        ("TCS", "TCS.NS"), ("Infosys", "INFY.NS"), ("HCL Tech", "HCLTECH.NS"), ("Wipro", "WIPRO.NS"),
        ("LTIMindtree", "LTIM.NS"), ("Tech Mahindra", "TECHM.NS"), ("Persistent Systems", "PERSISTENT.NS"),
        ("Coforge", "COFORGE.NS"), ("Tata Elxsi", "TATAELXSI.NS"), ("KPIT Technologies", "KPITTECH.NS")
    ],
    "NIFTY AUTO & EV": [
        ("Tata Motors", "TATAMOTORS.NS"), ("Mahindra & Mahindra", "M&M.NS"), ("Maruti Suzuki", "MARUTI.NS"),
        ("Bajaj Auto", "BAJAJ-AUTO.NS"), ("TVS Motor", "TVSMOTOR.NS"), ("Eicher Motors", "EICHERMOT.NS"),
        ("Ola Electric", "OLAELEC.NS"), ("Olectra Greentech", "OLECTRA.NS"), ("Exide Industries", "EXIDEIND.NS"),
        ("Amara Raja Energy", "ARE&M.NS"), ("Sona BLW", "SONACOMS.NS")
    ],
    "NIFTY GREEN ENERGY & POWER": [
        ("Tata Power", "TATAPOWER.NS"), ("Suzlon Energy", "SUZLON.NS"), ("IREDA", "IREDA.NS"),
        ("Adani Green Energy", "ADANIGREEN.NS"), ("Inox Wind", "INOXWIND.NS"), ("KPI Green Energy", "KPIGREEN.NS"),
        ("Waaree Energies", "WAAREE.NS"), ("Premier Energies", "PREMIERENE.NS"), ("NTPC Ltd", "NTPC.NS"),
        ("NHPC Ltd", "NHPC.NS"), ("JSW Energy", "JSWENERGY.NS")
    ],
    "NIFTY DEFENCE & RAILWAYS": [
        ("HAL (Hindustan Aero)", "HAL.NS"), ("Bharat Electronics", "BEL.NS"), ("Mazagon Dock", "MAZDOCK.NS"),
        ("Cochin Shipyard", "COCHINSHIP.NS"), ("Bharat Dynamics", "BDL.NS"), ("Solar Industries", "SOLARINDS.NS"),
        ("IRFC", "IRFC.NS"), ("RVNL", "RVNL.NS"), ("IRCTC", "IRCTC.NS"), ("Titagarh Rail", "TITAGARH.NS"),
        ("Jupiter Wagons", "JWL.NS"), ("RailTel", "RAILTEL.NS"), ("IRCON Intl", "IRCON.NS")
    ],
    "GLOBAL US TECH & MEGA CAPS": [
        ("Apple Inc.", "AAPL"), ("Microsoft Corp", "MSFT"), ("NVIDIA Corp", "NVDA"),
        ("Alphabet (Google)", "GOOGL"), ("Amazon.com", "AMZN"), ("Meta Platforms", "META"),
        ("Tesla Inc.", "TSLA"), ("Broadcom Inc.", "AVGO"), ("AMD", "AMD"),
        ("Taiwan Semi (TSMC)", "TSM"), ("Qualcomm", "QCOM"), ("Palantir Tech", "PLTR")
    ]
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

# --- TECHNICAL CALCULATIONS ---
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

# --- TOP TRADINGVIEW MULTI-INDEX & STOCK SCREENER ---
st.markdown(f"<div class='sec-header'>{get_txt('🔎 TradingView स्टॉक स्क्रीनर व इंडेक्स फ़िल्टर', 'TradingView Stock Screener & Index Filters')}</div>", unsafe_allow_html=True)

col_idx, col_stk = st.columns([1, 2])
with col_idx:
    selected_index = st.selectbox(
        get_txt("📂 इंडेक्स / सेक्टर चुनें:", "Select Index / Sector:"),
        list(INDEX_STOCKS_MAP.keys())
    )

stock_items = INDEX_STOCKS_MAP[selected_index]
stock_map = {f"{name} [{ticker}]": ticker for name, ticker in stock_items}
stock_options = list(stock_map.keys()) + ["➕ Type Any Custom Symbol (अन्य कोई भी भारतीय / US सिंबल)"]

with col_stk:
    selected_stock_display = st.selectbox(
        get_txt("🔎 स्टॉक चुनें या नाम टाइप करें (Type to Search Any Stock):", "Select or Type Stock Name:"),
        options=stock_options,
        index=0,
        help="बैंक ऑफ बड़ौदा, केनरा बैंक, टाटा या कोई भी नाम टाइप करें।"
    )

if selected_stock_display == "➕ Type Any Custom Symbol (अन्य कोई भी भारतीय / US सिंबल)":
    custom_sym = st.text_input(
        get_txt("स्टॉक सिंबल दर्ज करें (उदा. BANKBARODA.NS, CANBK.NS, NVDA):", "Enter Stock Symbol:"),
        value="BANKBARODA.NS"
    ).strip().upper()
    symbol = f"{custom_sym}.NS" if ("." not in custom_sym and not custom_sym.startswith("^") and len(custom_sym) > 4 and custom_sym.isalpha()) else custom_sym
else:
    symbol = stock_map[selected_stock_display]

# Duration & Range
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

# Screener Table Expander (Like TradingView Multi-stock table)
with st.expander(f"📊 {selected_index} - {get_txt('लाइव स्क्रीनर व AI बाय/सेल सिग्नल ग्रिड', 'Live Screener & Signals Grid')}", expanded=False):
    if st.button("⚡ रन स्क्रीनर स्कैन (Run Screener Scan)", key="run_screener"):
        with st.spinner("Scanning index stocks..."):
            screener_rows = []
            for s_name, s_ticker in stock_items:
                try:
                    s_t = yf.Ticker(s_ticker)
                    s_h = s_t.history(period="3mo")
                    if not s_h.empty and len(s_h) >= 15:
                        c_p = round(s_h["Close"].iloc[-1], 2)
                        p_c = round(s_h["Close"].iloc[-2], 2) if len(s_h) > 1 else c_p
                        chg_pct = round(((c_p - p_c) / p_c) * 100, 2)
                        r_val = round(calculate_rsi(s_h["Close"]).iloc[-1], 1)
                        
                        # AI Action
                        if r_val < 35: act = "🟢 BUY (Oversold)"
                        elif r_val > 70: act = "🔴 SELL (Overbought)"
                        elif chg_pct > 1.5 and r_val > 55: act = "🟢 STRONG BUY"
                        elif chg_pct < -1.5 and r_val < 45: act = "🔴 STRONG SELL"
                        else: act = "🟡 HOLD"

                        screener_rows.append({
                            "Symbol": s_ticker,
                            "Company Name": s_name,
                            "Price (CMP)": c_p,
                            "Change %": f"{chg_pct:+}%",
                            "RSI (14)": r_val,
                            "AI Action Signal": act
                        })
                except Exception:
                    continue
            if screener_rows:
                st.dataframe(pd.DataFrame(screener_rows), use_container_width=True)

# Fetch Stock Data
@st.cache_data(ttl=60)
def fetch_stock_payload(ticker_symbol, period_val, s_date, e_date):
    try:
        t = yf.Ticker(ticker_symbol)
        h = t.history(period=period_val) if period_val else t.history(start=s_date, end=e_date)
        max_h = t.history(period="max")
        info = t.info
        divs = t.dividends
        if h.empty:
            return None, None, None, None, None
        ath = max_h["High"].max() if not max_h.empty else None
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

    if df_hist is not None and stock_info is not None:
        cmp_price = stock_info.get("currentPrice") or stock_info.get("regularMarketPrice") or df_hist["Close"].iloc[-1]
        prev_close = stock_info.get("regularMarketPreviousClose") or cmp_price
        price_change = cmp_price - prev_close
        price_change_pct = (price_change / prev_close) * 100 if prev_close else 0.0

        high_52 = stock_info.get("fiftyTwoWeekHigh") or df_hist["High"].max()
        low_52 = stock_info.get("fiftyTwoWeekLow") or df_hist["Low"].min()
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
        sector = stock_info.get("sector", "N/A")
        industry = stock_info.get("industry", "N/A")

        # Technical Indicators Calculations
        df_hist["SMA_20"] = df_hist["Close"].rolling(20).mean()
        df_hist["SMA_50"] = df_hist["Close"].rolling(50).mean()
        df_hist["SMA_200"] = df_hist["Close"].rolling(200).mean()
        df_hist["Upper_BB"], df_hist["Lower_BB"], _ = calculate_bollinger_bands(df_hist["Close"])
        df_hist["RSI"] = calculate_rsi(df_hist["Close"])
        df_hist["MACD"], df_hist["MACD_Sig"] = calculate_macd(df_hist["Close"])

        latest_rsi = df_hist["RSI"].iloc[-1] if not np.isnan(df_hist["RSI"].iloc[-1]) else 50.0
        latest_macd = df_hist["MACD"].iloc[-1]
        latest_sig = df_hist["MACD_Sig"].iloc[-1]
        sma_50_val = df_hist["SMA_50"].iloc[-1] if not np.isnan(df_hist["SMA_50"].iloc[-1]) else cmp_price

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
        ai_action = "STRONG BUY 🚀" if win_prob >= 75 else ("BUY 📈" if win_prob >= 58 else ("HOLD ⚖️" if win_prob >= 45 else "SELL / AVOID 📉"))

        # Suggested Levels
        entry_lvl = round(cmp_price * 0.985, 2)
        sl_lvl = round(cmp_price * 0.94, 2)
        tgt_1 = round(cmp_price * 1.08, 2)
        tgt_2 = round(cmp_price * 1.15, 2)

        # Brokerage Ratings
        analyst_recom = stock_info.get("recommendationKey", "N/A").upper()
        target_mean = stock_info.get("targetMeanPrice", cmp_price * 1.12)

        intrinsic_val = calculate_intrinsic_value(eps, book_val) if eps and book_val else None

        # Dividend Analytics
        total_lifetime_div = df_div.sum() if not df_div.empty else 0.0
        yield_on_cost = (div_rate / buy_price * 100) if buy_price > 0 else None

        # Stock Header
        st.markdown("---")
        st.subheader(f"🏢 {long_name} ({symbol})")
        st.caption(f"Sector: **{sector}** | Industry: **{industry}** | Currency: **{currency}**")

        # 1. Price Action & ATH Range
        st.markdown(f"<div class='sec-header'>{get_txt('मूल्य एवं 52-सप्ताह/लाइफटाइम स्थिति', 'Price Action & ATH Range')}</div>", unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("CMP", f"{currency} {cmp_price:,.2f}", f"{price_change:+.2f} ({price_change_pct:+.2f}%)")
        m2.metric("52W High", f"{currency} {high_52:,.2f}" if high_52 else "N/A", f"{down_from_52w:.2f}% (from 52W High)", delta_color="inverse")
        m3.metric("52W Low", f"{currency} {low_52:,.2f}" if low_52 else "N/A")
        m4.metric("Lifetime High (ATH)", f"{currency} {ath:,.2f}" if ath else "N/A", f"{down_from_ath:.2f}% (from ATH)", delta_color="inverse")
        if intrinsic_val:
            m5.metric("Intrinsic Value", f"{currency} {intrinsic_val:,.2f}", f"{((intrinsic_val - cmp_price) / cmp_price) * 100:+.1f}% Margin")
        else:
            m5.metric("Intrinsic Value", "N/A")

        # 2. Valuation & Fundamentals
        st.markdown(f"<div class='sec-header'>{get_txt('वैल्युएशन एवं फंडामेंटल्स (P/E Multiples)', 'Valuation & Fundamentals')}</div>", unsafe_allow_html=True)
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Stock P/E", f"{company_pe:.2f}" if company_pe else "N/A")
        v2.metric("Industry P/E", str(industry_pe))
        v3.metric("P/B Ratio", f"{pb_ratio:.2f}" if pb_ratio else "N/A")
        v4.metric("EPS (TTM)", f"{currency} {eps:.2f}" if eps else "N/A")
        v5.metric("Dividend Yield (CMP)", f"{div_yield:.2f}%")

        # 3. Dividend Intelligence & Yield on Cost
        st.markdown(f"<div class='sec-header'>{get_txt('💰 डिविडेंड विश्लेषण एवं पूंजी यील्ड (Dividend Analytics & Yield on Cost)', 'Dividend Analytics & Yield on Cost')}</div>", unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.metric(get_txt("लाइफटाइम कुल डिविडेंड", "Lifetime Total Div"), f"{currency} {total_lifetime_div:,.2f}")
        d2.metric(get_txt("वार्षिक डिविडेंड दर (TTM)", "Annual Div Rate (TTM)"), f"{currency} {div_rate:,.2f}")
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
        st.markdown(f"<div class='sec-header'>{get_txt('💎 AI एक्सपर्ट रिपोर्ट व खरीद/बिक्री निर्णय (Buy/Sell Recommendation)', 'AI Expert Report & Signals')}</div>", unsafe_allow_html=True)

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
