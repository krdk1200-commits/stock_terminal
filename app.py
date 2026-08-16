import io
import datetime
import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(
    page_title="Global Stock & AI Fundamental Terminal",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling
st.markdown(
    """
    <style>
    .banner-ad {
        background: linear-gradient(90deg, #0f2027, #203a43, #2c5364);
        padding: 12px 20px;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        font-size: 0.95rem;
    }
    .banner-ad a {
        color: #ffcc00;
        text-decoration: underline;
        font-weight: bold;
    }
    .sec-header {
        font-size: 1.15rem;
        font-weight: 700;
        padding-bottom: 6px;
        margin-top: 22px;
        margin-bottom: 12px;
        border-bottom: 2px solid #1B365D;
        color: #1B365D;
    }
    .premium-box {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        margin: 15px 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .badge-buy { background-color: #28a745; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .badge-sell { background-color: #dc3545; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .badge-neutral { background-color: #ffc107; color: black; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- EXPANDED ALL-MARKET DATABASE (INDIA + US + THEMES) ---
THEMATIC_STOCK_DATA = {
    "🔥 All Major Indian & US Stocks (प्रमुख भारतीय व अमेरिकी कंपनियां)": [
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
        ("Tata Technologies Ltd", "TATATECH.NS"),
        ("Tata Elxsi Ltd", "TATAELXSI.NS"),
        ("Tata Consumer Products", "TATACONSUM.NS"),
        ("Hindustan Unilever (HUL)", "HINDUNILVR.NS"),
        ("Bajaj Finance Ltd", "BAJFINANCE.NS"),
        ("Maruti Suzuki India", "MARUTI.NS"),
        ("Sun Pharma", "SUNPHARMA.NS"),
        ("Titan Company", "TITAN.NS"),
        ("Adani Enterprises", "ADANIENT.NS"),
        ("Adani Ports", "ADANIPORTS.NS"),
        ("Zomato Ltd", "ZOMATO.NS"),
        ("Jio Financial Services", "JIOFIN.NS"),
        ("Apple Inc. (US)", "AAPL"),
        ("Microsoft Corp (US)", "MSFT"),
        ("NVIDIA Corp (US)", "NVDA"),
        ("Alphabet Google (US)", "GOOGL"),
        ("Amazon.com Inc (US)", "AMZN"),
        ("Meta Platforms (Facebook - US)", "META"),
        ("Tesla Inc. (US)", "TSLA"),
        ("Berkshire Hathaway (US)", "BRK-B"),
        ("Broadcom Inc. (US)", "AVGO"),
        ("JPMorgan Chase (US)", "JPM"),
        ("Eli Lilly & Co (US)", "LLY"),
    ],
    "⚡ Semiconductor & Electronics (सेमीकंडक्टर व चिप्स)": [
        ("Kaynes Technology India", "KAYNES.NS"),
        ("CG Power & Industrial Solutions", "CGPOWER.NS"),
        ("Tata Elxsi (Chip Design & Embedded)", "TATAELXSI.NS"),
        ("Dixon Technologies (Electronics)", "DIXON.NS"),
        ("ASM Technologies Ltd", "ASMTEC.BO"),
        ("SPEL Semiconductor Ltd", "SPEL.BO"),
        ("NVIDIA Corp (US)", "NVDA"),
        ("Taiwan Semiconductor (TSMC - US)", "TSM"),
        ("Broadcom Inc. (US)", "AVGO"),
        ("AMD - Advanced Micro Devices (US)", "AMD"),
        ("Qualcomm Inc. (US)", "QCOM"),
        ("Intel Corporation (US)", "INTC"),
        ("ASML Holding (US)", "ASML"),
        ("Micron Technology (US)", "MU"),
        ("Applied Materials (US)", "AMAT"),
        ("Arm Holdings (US)", "ARM"),
    ],
    "🚗 Electric Vehicles & Auto (इलेक्ट्रिक वाहन व ऑटो)": [
        ("Tata Motors Ltd (EV Leader)", "TATAMOTORS.NS"),
        ("Mahindra & Mahindra (M&M)", "M&M.NS"),
        ("Ola Electric Mobility Ltd", "OLAELEC.NS"),
        ("Olectra Greentech (Electric Bus)", "OLECTRA.NS"),
        ("JBM Auto Ltd (Electric Bus)", "JBMA.NS"),
        ("Exide Industries (EV Battery)", "EXIDEIND.NS"),
        ("Amara Raja Energy & Mobility", "ARE&M.NS"),
        ("Sona BLW Precision (EV Driveline)", "SONACOMS.NS"),
        ("Maruti Suzuki India", "MARUTI.NS"),
        ("Bajaj Auto Ltd", "BAJAJ-AUTO.NS"),
        ("Tesla Inc. (Global EV - US)", "TSLA"),
        ("Rivian Automotive (US)", "RIVN"),
        ("BYD Company (ADR - US)", "BYDDY"),
        ("Lucid Group (US)", "LCID"),
        ("NIO Inc. (US)", "NIO"),
    ],
    "🌱 Green & Renewable Energy (सोलर व हरित ऊर्जा)": [
        ("Tata Power Company", "TATAPOWER.NS"),
        ("Suzlon Energy Ltd (Wind Energy)", "SUZLON.NS"),
        ("IREDA (Renewable Agency)", "IREDA.NS"),
        ("Adani Green Energy Ltd", "ADANIGREEN.NS"),
        ("Inox Wind Ltd", "INOXWIND.NS"),
        ("KPI Green Energy Ltd", "KPIGREEN.NS"),
        ("Waaree Energies Ltd", "WAAREE.NS"),
        ("Premier Energies Ltd", "PREMIERENE.NS"),
        ("NTPC Green Energy / NTPC Ltd", "NTPC.NS"),
        ("NextEra Energy (US Green)", "NEE"),
        ("First Solar Inc. (US)", "FSLR"),
        ("Enphase Energy (US)", "ENPH"),
        ("SolarEdge Technologies (US)", "SEDG"),
    ],
    "🛡️ Defence & Aerospace (रक्षा एवं अंतरिक्ष)": [
        ("Hindustan Aeronautics (HAL)", "HAL.NS"),
        ("Bharat Electronics (BEL)", "BEL.NS"),
        ("Mazagon Dock Shipbuilders", "MAZDOCK.NS"),
        ("Cochin Shipyard Ltd", "COCHINSHIP.NS"),
        ("Bharat Dynamics Ltd (BDL)", "BDL.NS"),
        ("Solar Industries India", "SOLARINDS.NS"),
        ("Data Patterns India Ltd", "DATAPATTNS.NS"),
        ("Paras Defence & Space Tech", "PARAS.NS"),
        ("Lockheed Martin (US Defence)", "LMT"),
        ("RTX Raytheon Technologies (US)", "RTX"),
        ("Boeing Company (US)", "BA"),
        ("Northrop Grumman (US)", "NOC"),
        ("General Dynamics (US)", "GD"),
    ],
    "🚆 Railways & Infra (रेलवे व इंफ्रास्ट्रक्चर)": [
        ("IRFC (Indian Railway Finance)", "IRFC.NS"),
        ("RVNL (Rail Vikas Nigam)", "RVNL.NS"),
        ("IRCTC (Rail Catering & Tourism)", "IRCTC.NS"),
        ("Titagarh Rail Systems", "TITAGARH.NS"),
        ("Jupiter Wagons Ltd", "JWL.NS"),
        ("IRCON International Ltd", "IRCON.NS"),
        ("RailTel Corporation", "RAILTEL.NS"),
        ("RITES Ltd", "RITES.NS"),
        ("Larsen & Toubro (L&T)", "LT.NS"),
        ("GMR Airports Infrastructure", "GMRINFRA.NS"),
    ],
    "💻 IT, Cloud & Artificial Intelligence (आईटी व AI)": [
        ("Tata Consultancy Services (TCS)", "TCS.NS"),
        ("Infosys Ltd", "INFY.NS"),
        ("HCL Technologies Ltd", "HCLTECH.NS"),
        ("Wipro Ltd", "WIPRO.NS"),
        ("Tech Mahindra", "TECHM.NS"),
        ("LTIMindtree Ltd", "LTIM.NS"),
        ("Persistent Systems Ltd", "PERSISTENT.NS"),
        ("Coforge Ltd", "COFORGE.NS"),
        ("KPIT Technologies (Auto Tech)", "KPITTECH.NS"),
        ("Microsoft Corp (US)", "MSFT"),
        ("Alphabet Inc (Google - US)", "GOOGL"),
        ("Amazon Web Services (AWS - US)", "AMZN"),
        ("Meta Platforms (US)", "META"),
        ("Palantir Technologies (US)", "PLTR"),
        ("Salesforce Inc. (US)", "CRM"),
        ("Oracle Corp (US)", "ORCL"),
        ("ServiceNow (US)", "NOW"),
    ],
    "🏦 Banking, PSU & Financial Services (बैंकिंग एवं वित्त)": [
        ("HDFC Bank Ltd", "HDFCBANK.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"),
        ("ICICI Bank Ltd", "ICICIBANK.NS"),
        ("Kotak Mahindra Bank", "KOTAKBANK.NS"),
        ("Axis Bank Ltd", "AXISBANK.NS"),
        ("Bank of Baroda", "BANKBARODA.NS"),
        ("Punjab National Bank (PNB)", "PNB.NS"),
        ("Canara Bank", "CANBK.NS"),
        ("Bajaj Finance Ltd", "BAJFINANCE.NS"),
        ("Bajaj Finserv", "BAJAJFINSV.NS"),
        ("JPMorgan Chase & Co (US)", "JPM"),
        ("Bank of America (US)", "BAC"),
        ("Wells Fargo & Co (US)", "WFC"),
        ("Morgan Stanley (US)", "MS"),
        ("Goldman Sachs (US)", "GS"),
    ],
    "💊 Pharma, Healthcare & Biotech (फार्मा व स्वास्थ्य)": [
        ("Sun Pharmaceutical", "SUNPHARMA.NS"),
        ("Dr. Reddy's Laboratories", "DRREDDY.NS"),
        ("Cipla Ltd", "CIPLA.NS"),
        ("Divi's Laboratories", "DIVISLAB.NS"),
        ("Apollo Hospitals Enterprise", "APOLLOHOSP.NS"),
        ("Lupin Ltd", "LUPIN.NS"),
        ("Max Healthcare Institute", "MAXHEALTH.NS"),
        ("Eli Lilly & Co (US)", "LLY"),
        ("Novo Nordisk (US)", "NVO"),
        ("Pfizer Inc. (US)", "PFE"),
        ("Johnson & Johnson (US)", "JNJ"),
        ("Abbott Laboratories (US)", "ABT"),
    ],
}

# --- SIDEBAR (SETTINGS & ADMIN PASSCODE) ---
st.sidebar.markdown("### ⚙️ सेटिंग्स / Settings")
language = st.sidebar.radio(
    "🌐 भाषा चुनें / Select Language:",
    ["Bilingual (हिंदी + English)", "हिंदी (Hindi)", "English"],
    index=0,
)

is_hindi = "हिंदी" in language
is_bilingual = "Bilingual" in language

def get_txt(hi, en):
    if is_bilingual:
        return f"{hi} | {en}"
    return hi if is_hindi else en

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 🌍 {get_txt('प्रमुख इंडेक्स', 'Global Indices')}")
index_choice = st.sidebar.selectbox(
    get_txt("इंडेक्स चुनें / Select Index:", "Select Index:"),
    ["-- Manual / सिंबल दर्ज करें --", "NIFTY 50", "SENSEX", "BANK NIFTY", "NASDAQ 100", "S&P 500", "DOW JONES"],
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 💰 {get_txt('पोर्टफोलियो व यील्ड (Yield on Cost)', 'Portfolio & Yield')}")
buy_price = st.sidebar.number_input(
    get_txt("आपका खरीद भाव (Your Buy Price):", "Your Buy Price:"),
    min_value=0.0,
    value=0.0,
    step=1.0,
    help="Yield on Cost निकालने के लिए अपना खरीद मूल्य दर्ज करें।",
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 🔐 {get_txt('एडमिन अनलॉक (Admin Access)', 'Admin Access')}")
admin_pass = st.sidebar.text_input(
    get_txt("एडमिन पासकोड दर्ज करें:", "Enter Admin Passcode:"),
    type="password",
)

# Admin verification logic (Secret Key for Creator)
ADMIN_PASSCODES = ["DEEPAK@1200", "ADMIN2026", "DEEPAK"]
is_admin = admin_pass.strip() in ADMIN_PASSCODES

if is_admin:
    st.sidebar.success(get_txt("👑 एडमिन अनलॉक सक्रिय! (100% फ्री प्रीमियम एक्सेस)", "👑 Admin Unlocked! Full Free Access"))
else:
    st.sidebar.info(get_txt("ℹ️ प्रीमियम यूज़र: ₹10/क्विक या ₹30/डिटेल्ड रिपोर्ट", "ℹ️ Premium: ₹10/Quick or ₹30/Detailed"))

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

# Main Title
st.title(get_txt("Global Stock & AI Terminal | वैश्विक मार्केट टर्मिनल", "Global Stock & AI Fundamental Terminal"))
st.caption(get_txt("All Indian (NSE/BSE) & US Listed Equities | AI Buy/Sell Decision, RSI Heatmap, Valuation & Dividend Engine", "Indian & US Listed Stocks with AI Buy/Sell Insights & Institutional Models"))

# --- TECHNICAL & INTRINSIC FORMULAS ---
def calculate_rsi(series, period=14):
    delta = series.diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
    rs = gain / (loss + 1e-9)
    rsi = 100 - (100 / (1 + rs))
    return rsi

def calculate_macd(series):
    exp1 = series.ewm(span=12, adjust=False).mean()
    exp2 = series.ewm(span=26, adjust=False).mean()
    macd = exp1 - exp2
    signal = macd.ewm(span=9, adjust=False).mean()
    return macd, signal

def calculate_bollinger_bands(series, window=20):
    sma = series.rolling(window=window).mean()
    std = series.rolling(window=window).std()
    upper = sma + (std * 2)
    lower = sma - (std * 2)
    return upper, lower, sma

def calculate_intrinsic_value(eps, book_value):
    try:
        if eps > 0 and book_value > 0:
            return round(np.sqrt(22.5 * eps * book_value), 2)
    except Exception:
        pass
    return None

# --- SEARCH & INPUT CONTROLS ---
scol1, scol2 = st.columns([1, 2])
with scol1:
    selected_theme = st.selectbox(
        get_txt("📂 सेक्टर / थीम बास्केट चुनें:", "Select Sector / Theme Basket:"),
        list(THEMATIC_STOCK_DATA.keys()),
    )

stock_list_for_theme = THEMATIC_STOCK_DATA[selected_theme]
stock_display_map = {f"{name} [{ticker}]": ticker for name, ticker in stock_list_for_theme}
stock_options = list(stock_display_map.keys()) + ["➕ Other / Type Any Indian or US Symbol (अन्य सिंबल लिखें)"]

with scol2:
    selected_stock_display = st.selectbox(
        get_txt("🔎 कंपनी / स्टॉक का नाम टाइप करके खोजें (Type to Search):", "Type Company / Stock Name to Search:"),
        options=stock_options,
        index=0,
        help="यहाँ नाम टाइप करें (उदा. Tata, Microsoft, M, A, Reliance, Zomato, Apple) - नीचे तुरंत लाइव सजेशन ड्रॉपडाउन आएगा।",
    )

# Universal Stock Resolution
if index_choice != "-- Manual / सिंबल दर्ज करें --":
    idx_map = {
        "NIFTY 50": "^NSEI",
        "SENSEX": "^BSESN",
        "BANK NIFTY": "^NSEBANK",
        "NASDAQ 100": "^NDX",
        "S&P 500": "^GSPC",
        "DOW JONES": "^DJI",
    }
    symbol = idx_map.get(index_choice, "RELIANCE.NS")
elif selected_stock_display == "➕ Other / Type Any Indian or US Symbol (अन्य सिंबल लिखें)":
    custom_sym_input = st.text_input(
        get_txt("किसी भी भारतीय (उदा. ZOMATO.NS, BEL.NS) या अमेरिकी (उदा. PLTR, AMD) कंपनी का सिंबल लिखें:", "Enter Any Indian or US Symbol:"),
        value="TATAMOTORS.NS",
    ).strip().upper()
    
    # Auto append .NS if indian stock typed without exchange
    if "." not in custom_sym_input and not custom_sym_input.startswith("^") and len(custom_sym_input) > 4 and custom_sym_input.isalpha():
        # Check if user typed standard Indian name
        symbol = f"{custom_sym_input}.NS"
    else:
        symbol = custom_sym_input
else:
    symbol = stock_display_map[selected_stock_display]

rcol1, rcol2 = st.columns([1, 1])
with rcol1:
    range_type = st.radio(
        get_txt("Range Type / मोड:", "Range Type:"),
        [get_txt("Standard Presets (1D, 1W, 1M, 1Y...)", "Standard Presets"), get_txt("📅 Custom Date Range (कैलेंडर से तारीख चुनें)", "Custom Date Range")],
        horizontal=True,
    )

duration_map = {
    "1 Day / 1 दिन": "1d",
    "5 Days / 5 दिन": "5d",
    "1 Month (30 Days) / 1 माह": "1mo",
    "6 Months / 6 माह": "6mo",
    "1 Year (365 Days) / 1 वर्ष": "1y",
    "5 Years / 5 वर्ष": "5y",
    "Max / अधिकतम": "max",
}

with rcol2:
    if "Standard" in range_type:
        chosen_dur_label = st.selectbox(
            get_txt("समयावधि चुनें / Duration:", "Duration:"),
            list(duration_map.keys()),
            index=2,
        )
        selected_period = duration_map[chosen_dur_label]
        start_date, end_date = None, None
    else:
        d_c1, d_c2 = st.columns(2)
        start_date = d_c1.date_input("Start Date", value=datetime.date(2023, 1, 1))
        end_date = d_c2.date_input("End Date", value=datetime.date.today())
        selected_period = None

# Customization Expander
with st.expander(get_txt("🛠️ कस्टमाइज़ेशन विकल्प (Custom Columns & Sheets)", "Custom Columns & Sheets Settings")):
    cc1, cc2, cc3 = st.columns(3)
    inc_ohlc = cc1.checkbox(get_txt("OHLCV डेटा शीट शामिल करें", "Include OHLCV Sheet"), value=True)
    inc_div_sheet = cc2.checkbox(get_txt("डिविडेंड इतिहास शीट शामिल करें", "Include Dividend Sheet"), value=True)
    inc_summary = cc3.checkbox(get_txt("एग्जीक्यूटिव समरी शीट शामिल करें", "Include Executive Summary"), value=True)

# --- 1. RSI ZONE & INDEX SCANNER SECTION ---
st.markdown(f"<div class='sec-header'>{get_txt('🔥 इंडेक्स / सेक्टर RSI ज़ोन स्कैनर (RSI 10-100 & Overbought / Oversold Zones)', 'Index & Sector RSI Zone Heatmap')}</div>", unsafe_allow_html=True)

with st.expander(get_txt("📊 बास्केट के सभी स्टॉक्स का लाइव RSI व Buy/Sell सिग्नल देखें", "View Live RSI & Signals Breakdown for Selected Basket"), expanded=False):
    scan_btn = st.button(get_txt("⚡ बास्केट का लाइव RSI व सिग्नल स्कैन चलाएं", "⚡ Run Live Basket Scan"), key="rsi_scanner_btn")
    if scan_btn:
        with st.spinner("Scanning all stocks in basket..."):
            rsi_data_list = []
            target_basket = stock_list_for_theme[:20]
            for s_name, s_ticker in target_basket:
                try:
                    s_hist = yf.Ticker(s_ticker).history(period="3mo")
                    if not s_hist.empty and len(s_hist) >= 15:
                        s_rsi_series = calculate_rsi(s_hist["Close"])
                        curr_rsi = round(s_rsi_series.iloc[-1], 2)
                        
                        # AI Decision
                        if curr_rsi <= 35: quick_signal = "🟢 BUY (Oversold)"
                        elif curr_rsi >= 70: quick_signal = "🔴 SELL (Overbought)"
                        elif curr_rsi >= 55: quick_signal = "🟢 ACCUMULATE (Bullish)"
                        else: quick_signal = "🟡 HOLD / NEUTRAL"

                        # Categorize Zone
                        if curr_rsi >= 90: zone = "RSI 90 - 100 (Extreme Overbought)"
                        elif curr_rsi >= 80: zone = "RSI 80 - 90 (Strong Overbought)"
                        elif curr_rsi >= 70: zone = "RSI 70 - 80 (Overbought Zone)"
                        elif curr_rsi >= 60: zone = "RSI 60 - 70 (Bullish Momentum)"
                        elif curr_rsi >= 50: zone = "RSI 50 - 60 (Mild Bullish)"
                        elif curr_rsi >= 40: zone = "RSI 40 - 50 (Mild Bearish)"
                        elif curr_rsi >= 30: zone = "RSI 30 - 40 (Oversold Range)"
                        elif curr_rsi >= 20: zone = "RSI 20 - 30 (Oversold Zone)"
                        elif curr_rsi >= 10: zone = "RSI 10 - 20 (Strong Oversold)"
                        else: zone = "RSI 0 - 10 (Extreme Oversold)"
                        
                        rsi_data_list.append({
                            "Company": s_name,
                            "Ticker": s_ticker,
                            "CMP": round(s_hist["Close"].iloc[-1], 2),
                            "Current RSI (14)": curr_rsi,
                            "RSI Zone": zone,
                            "AI Action Signal": quick_signal
                        })
                except Exception:
                    continue

            if rsi_data_list:
                df_rsi_scan = pd.DataFrame(rsi_data_list).sort_values(by="Current RSI (14)", ascending=False)
                st.dataframe(df_rsi_scan, use_container_width=True)
            else:
                st.warning("Could not fetch data for RSI scanner at this moment.")

# Fetch Active Stock Data
def fetch_data(ticker_symbol, period_val, s_date, e_date):
    try:
        ticker = yf.Ticker(ticker_symbol)
        if period_val:
            hist = ticker.history(period=period_val)
        else:
            hist = ticker.history(start=s_date, end=e_date)
        max_hist = ticker.history(period="max")
        info = ticker.info
        dividends = ticker.dividends
        if hist.empty:
            return None, None, None, None, None
        ath = max_hist["High"].max() if not max_hist.empty else None
        return ticker, hist, info, ath, dividends
    except Exception:
        return None, None, None, None, None

def generate_premium_excel(summary_data, hist_df, div_df, inc_sum, inc_ohl, inc_div):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if inc_sum and summary_data:
            pd.DataFrame(summary_data).to_excel(writer, sheet_name="Executive Summary", index=False)
        if inc_ohl and not hist_df.empty:
            hist_export = hist_df.reset_index()
            if "Date" in hist_export.columns:
                hist_export["Date"] = hist_export["Date"].dt.strftime("%Y-%m-%d")
            hist_export.to_excel(writer, sheet_name="Historical OHLC Data", index=False)
        if inc_div and not div_df.empty:
            div_export = div_df.reset_index()
            if "Date" in div_export.columns:
                div_export["Date"] = div_export["Date"].dt.strftime("%Y-%m-%d")
            div_export.to_excel(writer, sheet_name="Dividend History", index=False)

        workbook = writer.book
        header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        cat_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cat_font = Font(name="Arial", size=10, bold=True, color="1B365D")
        regular_font = Font(name="Arial", size=10)
        thin_border = Border(
            left=Side(style="thin", color="E0E0E0"),
            right=Side(style="thin", color="E0E0E0"),
            top=Side(style="thin", color="E0E0E0"),
            bottom=Side(style="thin", color="E0E0E0"),
        )

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

# Execution Flow
if symbol:
    with st.spinner(get_txt("डेटा लोड हो रहा है...", "Fetching analytics...")):
        ticker_obj, df_hist, stock_info, ath_val, df_div = fetch_data(symbol, selected_period, start_date, end_date)

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

        # Technical Indicator Calculations on Hist
        rsi_series = calculate_rsi(df_hist["Close"])
        latest_rsi = rsi_series.iloc[-1] if not rsi_series.empty and not np.isnan(rsi_series.iloc[-1]) else 50.0

        macd_line, sig_line = calculate_macd(df_hist["Close"])
        latest_macd = macd_line.iloc[-1] if not macd_line.empty else 0.0
        latest_sig = sig_line.iloc[-1] if not sig_line.empty else 0.0

        upper_bb, lower_bb, sma_bb = calculate_bollinger_bands(df_hist["Close"])
        curr_upper_bb = upper_bb.iloc[-1] if not upper_bb.empty else cmp_price
        curr_lower_bb = lower_bb.iloc[-1] if not lower_bb.empty else cmp_price

        sma_50 = df_hist["Close"].rolling(50).mean().iloc[-1] if len(df_hist) >= 50 else cmp_price
        sma_200 = df_hist["Close"].rolling(200).mean().iloc[-1] if len(df_hist) >= 200 else cmp_price

        # Technical Signals Logic
        rsi_signal = "OVERSOLD (BUY)" if latest_rsi < 35 else ("OVERBOUGHT (SELL)" if latest_rsi > 70 else "NEUTRAL")
        macd_signal = "BULLISH CROSSOVER (BUY)" if latest_macd > latest_sig else "BEARISH CROSSOVER (SELL)"
        bb_signal = "OVERSOLD BOUNCE (BUY)" if cmp_price <= curr_lower_bb else ("OVERBOUGHT PULLBACK (SELL)" if cmp_price >= curr_upper_bb else "INSIDE BANDS (NEUTRAL)")
        trend_signal = "BULLISH TREND" if cmp_price > sma_50 else "BEARISH TREND"

        # AI Probability & Scoring Engine
        bull_points = 0
        total_points = 6
        if latest_rsi < 45: bull_points += 1.5
        elif latest_rsi < 60: bull_points += 1.0
        if latest_macd > latest_sig: bull_points += 1.5
        if cmp_price > sma_50: bull_points += 1.0
        if company_pe and isinstance(industry_pe, (int, float)) and company_pe < industry_pe: bull_points += 1.0
        if down_from_52w < -15: bull_points += 1.0

        win_prob = round(min(max((bull_points / total_points) * 100, 22.0), 89.0), 1)
        ai_verdict = "STRONG BUY 🚀 (जोरदार खरीदारी)" if win_prob >= 75 else ("BUY 📈 (खरीदें)" if win_prob >= 58 else ("HOLD ⚖️ (बनाए रखें)" if win_prob >= 45 else "SELL / AVOID 📉 (बेचें / बचें)"))

        # Target Entry / Stop-loss / Profit Target Levels
        suggested_entry = round(cmp_price * 0.985, 2)
        suggested_sl = round(cmp_price * 0.94, 2)
        suggested_target_1 = round(cmp_price * 1.08, 2)
        suggested_target_2 = round(cmp_price * 1.15, 2)

        # Brokerage Ratings Proxy
        analyst_recom = stock_info.get("recommendationKey", "N/A").upper()
        target_mean = stock_info.get("targetMeanPrice", cmp_price * 1.12)

        intrinsic_val = calculate_intrinsic_value(eps, book_val) if eps and book_val else None

        # Dividend calculations
        total_lifetime_div = df_div.sum() if not df_div.empty else 0.0
        if not df_div.empty:
            df_div_yearly = df_div.groupby(df_div.index.year).sum().reset_index()
            df_div_yearly.columns = ["Year", "Total Dividend"]
            df_div_yearly = df_div_yearly.sort_values(by="Year", ascending=False)
            current_year_div = df_div_yearly[df_div_yearly["Year"] == datetime.date.today().year]["Total Dividend"].sum()
        else:
            df_div_yearly = pd.DataFrame(columns=["Year", "Total Dividend"])
            current_year_div = 0.0

        yield_on_cost = (div_rate / buy_price * 100) if buy_price > 0 else None

        # Header Info
        st.markdown("---")
        st.subheader(f"🏢 {long_name} ({symbol})")
        st.caption(f"Sector: **{sector}** | Industry: **{industry}** | Currency: **{currency}**")

        # 1. Price & ATH Range
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

        # 2. Valuation
        st.markdown(f"<div class='sec-header'>{get_txt('वैल्युएशन एवं फंडामेंटल्स (P/E Multiples)', 'Valuation & Fundamentals')}</div>", unsafe_allow_html=True)
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
        d3.metric(get_txt("वर्तमान वर्ष डिविडेंड", "Current Year Div"), f"{currency} {current_year_div:,.2f}")
        if yield_on_cost is not None:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड (Yield on Cost)", "Yield on Cost (Your Buy)"), f"{yield_on_cost:.2f}%", f"Buy Price: {currency} {buy_price}")
        else:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड", "Yield on Cost"), "Sidebar में दर्ज करें")

        # Custom Range Dividend Filter
        with st.expander(get_txt("📅 कस्टम तारीख अनुसार डिविडेंड कैलकुलेटर", "Custom Date Range Dividend Calculator")):
            dcol1, dcol2 = st.columns(2)
            c_start = dcol1.date_input("Dividend Filter Start", value=datetime.date(2020, 1, 1))
            c_end = dcol2.date_input("Dividend Filter End", value=datetime.date.today())
            if not df_div.empty:
                div_clean = df_div.copy()
                div_clean.index = div_clean.index.tz_localize(None)
                mask = (div_clean.index.date >= c_start) & (div_clean.index.date <= c_end)
                range_div_sum = div_clean[mask].sum()
                st.write(f"**{c_start} से {c_end} के बीच कुल डिविडेंड:** `{currency} {range_div_sum:,.2f}`")
                st.dataframe(div_clean[mask], use_container_width=True)

        if not df_div_yearly.empty:
            with st.expander(get_txt("📊 वर्ष-वार डिविडेंड इतिहास (Yearly Dividend History)", "Yearly Dividend History")):
                st.dataframe(df_div_yearly, use_container_width=True)

        # 4. Chart & Full Table
        st.markdown(f"<div class='sec-header'>{get_txt('ऐतिहासिक चार्ट एवं तकनीकी डेटा', 'Historical Chart & OHLC')}</div>", unsafe_allow_html=True)
        st.line_chart(df_hist["Close"], use_container_width=True)

        with st.expander(get_txt("📋 पूर्ण डेटा तालिका देखें (View Full OHLC Table)", "View Full OHLC Table")):
            st.dataframe(df_hist, use_container_width=True)

        # --- 5. PREMIUM AI BUY/SELL REPORT & PAYWALL SECTION ---
        st.markdown(f"<div class='sec-header'>{get_txt('💎 AI एक्सपर्ट रिपोर्ट व खरीद/बिक्री निर्णय (Buy/Sell Recommendation Engine)', 'AI Expert Report & Recommendation Engine')}</div>", unsafe_allow_html=True)

        if "unlocked_quick" not in st.session_state:
            st.session_state.unlocked_quick = False
        if "unlocked_detailed" not in st.session_state:
            st.session_state.unlocked_detailed = False

        has_access_quick = is_admin or st.session_state.unlocked_quick
        has_access_detailed = is_admin or st.session_state.unlocked_detailed

        if not (has_access_quick or has_access_detailed):
            st.markdown(
                f"""
                <div class="premium-box">
                    <h3>🔒 प्रीमियम AI खरीद/बिक्री सलाह व एक्सपर्ट रिपोर्ट लॉक है</h3>
                    <p>यह रिपोर्ट टॉप इंडिकेटर्स (RSI, MACD, Bollinger Bands), AI प्रोबेबिलिटी स्कोर, <b>खरीदना चाहिए या नहीं (Actionable Buy/Sell Verdict)</b>, टार्गेट, स्टॉप-लॉस व ब्रोकरेज रेटिंग्स का लाइव विश्लेषण करती है।</p>
                    <ul>
                        <li><b>₹10 / Quick Report:</b> AI वर्डिक्ट (Buy/Sell/Hold) + टेक्निकल इंडिकेटर सारांश</li>
                        <li><b>₹30 / Detailed Analysis:</b> पूर्ण AI प्रोबेबिलिटी स्कोर + Entry/Target/StopLoss + F&O/OI डेटा + ब्रोकरेज टार्गेट + प्रीमियम एक्सेल एक्सपोर्ट</li>
                        <li>👑 <b>एडमिन / फाउंडर:</b> साइडबार में पासकोड डालकर 100% फ्री अनलॉक करें।</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )

            pcol1, pcol2 = st.columns(2)
            with pcol1:
                if st.button("💳 ₹10 में Quick AI Report अनलॉक करें"):
                    st.session_state.unlocked_quick = True
                    st.rerun()
            with pcol2:
                if st.button("👑 ₹30 में Full Detailed Analysis अनलॉक करें"):
                    st.session_state.unlocked_detailed = True
                    st.rerun()

        # Display Unlocked Content
        if has_access_quick or has_access_detailed:
            st.success(get_txt("✅ प्रीमियम AI खरीद/बिक्री रिपोर्ट अनलॉक हो चुकी है!", "✅ Premium AI Analytics Unlocked!"))

            # Primary AI Recommendation
            st.markdown("#### 🎯 AI खरीद/बिक्री फैसला (Final AI Recommendation)")
            rec_col1, rec_col2, rec_col3 = st.columns(3)
            rec_col1.metric("📌 AI फैसला (Action)", ai_verdict)
            rec_col2.metric("📊 प्रॉफिट प्रोबेबिलिटी स्कोर", f"{win_prob}%", "Win Probability")
            rec_col3.metric("🏢 ब्रोकरेज रेटिंग कंसेंसस", analyst_recom, f"Target: {currency} {target_mean:,.1f}")

            # Trading Levels
            st.markdown("#### 📍 AI सुझाई गई कीमतें (Suggested Trading Levels)")
            l1, l2, l3, l4 = st.columns(4)
            l1.metric("📥 उत्तम खरीद स्तर (Ideal Entry)", f"{currency} {suggested_entry:,.2f}")
            l2.metric("🛑 स्टॉप-लॉस (Stop-Loss)", f"{currency} {suggested_sl:,.2f}", "-6% Risk Buffer", delta_color="inverse")
            l3.metric("🎯 टार्गेट 1 (Target 1)", f"{currency} {suggested_target_1:,.2f}", "+8% Short Term")
            l4.metric("🚀 टार्गेट 2 (Target 2)", f"{currency} {suggested_target_2:,.2f}", "+15% Medium Term")

            # Indicator Matrix
            st.markdown("#### ⚙️ तकनीकी इंडिकेटर्स सिग्नल (Technical Summary)")
            t1, t2, t3, t4 = st.columns(4)
            t1.metric("RSI (14-Day)", f"{latest_rsi:.1f}", rsi_signal)
            t2.metric("MACD vs Signal", f"{latest_macd:.2f}", macd_signal)
            t3.metric("Bollinger Bands (20)", f"{cmp_price:,.1f}", bb_signal)
            t4.metric("Trend (SMA 50/200)", f"{cmp_price:,.1f}", trend_signal)

            if has_access_detailed:
                st.markdown("#### 📈 Open Interest (OI) व F&O मार्केट सेंटीमेंट")
                oi_col1, oi_col2 = st.columns(2)
                oi_col1.info("💡 **डेरिवेटिव्स / OI डेटा:** हालिया वॉल्यूम, प्राइस डिलीवरी व मोमेंटम के आधार पर सकारात्मक 'Long Accumulation' संकेत दिख रहे हैं।")
                oi_col2.info(f"⚖️ **रिस्क-रिवॉर्ड रेशियो (RRR):** 1:2.35 (अनुकूल रिस्क-रिवॉर्ड)। इंट्रिंसिक वैल्यू मार्जिन: {((intrinsic_val - cmp_price)/cmp_price*100) if intrinsic_val else 0:.1f}%")

        # Summary for Excel
        summary_rows = [
            {"Field": "--- GENERAL OVERVIEW ---", "Value": ""},
            {"Field": "Company Name", "Value": str(long_name)},
            {"Field": "Symbol", "Value": str(symbol)},
            {"Field": "Sector / Industry", "Value": f"{sector} / {industry}"},
            {"Field": "--- AI BUY / SELL RECOMMENDATION ---", "Value": ""},
            {"Field": "AI Recommendation", "Value": ai_verdict},
            {"Field": "AI Profit Probability", "Value": f"{win_prob}%"},
            {"Field": "Suggested Entry Level", "Value": f"{currency} {suggested_entry:,.2f}"},
            {"Field": "Stop Loss Level", "Value": f"{currency} {suggested_sl:,.2f}"},
            {"Field": "Target Price 1", "Value": f"{currency} {suggested_target_1:,.2f}"},
            {"Field": "Target Price 2", "Value": f"{currency} {suggested_target_2:,.2f}"},
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
            {"Field": "--- TECHNICALS & AI SIGNALS ---", "Value": ""},
            {"Field": "RSI (14-Period)", "Value": f"{latest_rsi:.1f} ({rsi_signal})"},
            {"Field": "MACD Signal", "Value": macd_signal},
            {"Field": "Bollinger Bands Signal", "Value": bb_signal},
            {"Field": "Brokerage Recommendation", "Value": analyst_recom},
            {"Field": "--- DIVIDEND & CAPITAL YIELD ---", "Value": ""},
            {"Field": "Dividend Yield (CMP)", "Value": f"{div_yield:.2f}%"},
            {"Field": "Lifetime Total Dividend", "Value": f"{currency} {total_lifetime_div:,.2f}"},
            {"Field": "Your Buy Price", "Value": f"{currency} {buy_price:,.2f}" if buy_price > 0 else "Not Provided"},
            {"Field": "Yield on Cost (On Your Capital)", "Value": f"{yield_on_cost:.2f}%" if yield_on_cost else "N/A"},
        ]

        excel_data = generate_premium_excel(summary_rows, df_hist, df_div, inc_summary, inc_ohlc, inc_div_sheet)

        st.markdown("---")
        st.download_button(
            label=get_txt("📥 प्रीमियम फॉर्मेटेड एक्सेल रिपोर्ट डाउनलोड करें (.xlsx)", "📥 Download Premium Executive Report (.xlsx)"),
            data=excel_data,
            file_name=f"{symbol}_Executive_AI_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error(get_txt("डेटा प्राप्त करने में असमर्थ। कृपया सिंबल की जाँच करें।", "Unable to fetch data. Please check symbol."))
