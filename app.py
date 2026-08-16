import datetime
import io
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import yfinance as yf

# 1. Page Configuration
st.set_page_config(
    page_title="TradingView Pro | Global Stock, F&O & Fundamental AI Terminal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Custom Styling
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
    .fund-sound { background-color: #d4edda; color: #155724; padding: 6px 12px; border-radius: 6px; font-weight: bold; border-left: 5px solid #28a745; }
    .fund-mod { background-color: #fff3cd; color: #856404; padding: 6px 12px; border-radius: 6px; font-weight: bold; border-left: 5px solid #ffc107; }
    .fund-weak { background-color: #f8d7da; color: #721c24; padding: 6px 12px; border-radius: 6px; font-weight: bold; border-left: 5px solid #dc3545; }
    </style>
    """,
    unsafe_allow_html=True,
)

# 3. Master Database of Stocks & Thematic Baskets
ALL_SECTORS_AND_INDICES = {
    "🔥 All Stocks (सभी प्रमुख भारतीय व अमेरिकी स्टॉक्स)": [
        ("Bank of India", "BANKINDIA.NS"), ("Bank of Baroda", "BANKBARODA.NS"), ("Canara Bank", "CANBK.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"), ("Punjab National Bank (PNB)", "PNB.NS"),
        ("Union Bank of India", "UNIONBANK.NS"), ("Indian Bank", "INDIANB.NS"), ("Central Bank of India", "CENTRALBK.NS"),
        ("Reliance Industries (RIL)", "RELIANCE.NS"), ("Tata Consultancy Services (TCS)", "TCS.NS"),
        ("Tata Motors Ltd", "TATAMOTORS.NS"), ("Tata Steel Ltd", "TATASTEEL.NS"), ("Tata Power Co Ltd", "TATAPOWER.NS"),
        ("Tata Technologies", "TATATECH.NS"), ("Tata Elxsi Ltd", "TATAELXSI.NS"), ("HDFC Bank Ltd", "HDFCBANK.NS"),
        ("ICICI Bank Ltd", "ICICIBANK.NS"), ("Infosys Ltd", "INFY.NS"), ("Bharti Airtel", "BHARTIARTL.NS"),
        ("ITC Ltd", "ITC.NS"), ("Larsen & Toubro (L&T)", "LT.NS"), ("Zomato Ltd", "ZOMATO.NS"),
        ("Jio Financial Services", "JIOFIN.NS"), ("Apple Inc (US)", "AAPL"), ("Microsoft Corp (US)", "MSFT"),
        ("NVIDIA Corp (US)", "NVDA"), ("Tesla Inc (US)", "TSLA")
    ],
    "⚡ Semiconductor & Electronics (सेमीकंडक्टर व चिप्स)": [
        ("Kaynes Technology India", "KAYNES.NS"), ("CG Power & Industrial", "CGPOWER.NS"),
        ("Tata Elxsi (Chip Design/AI)", "TATAELXSI.NS"), ("Dixon Technologies", "DIXON.NS"),
        ("ASM Technologies Ltd", "ASMTEC.BO"), ("SPEL Semiconductor Ltd", "SPEL.BO"),
        ("NVIDIA Corp (US)", "NVDA"), ("Taiwan Semiconductor (TSMC - US)", "TSM"),
        ("Broadcom Inc (US)", "AVGO"), ("AMD (US)", "AMD"), ("Qualcomm (US)", "QCOM"),
        ("Intel Corporation (US)", "INTC"), ("ASML Holding (US)", "ASML"), ("Micron Technology (US)", "MU")
    ],
    "🚆 Railways, Defence & Aerospace (रेलवे एवं रक्षा क्षेत्र)": [
        ("HAL (Hindustan Aeronautics)", "HAL.NS"), ("Bharat Electronics (BEL)", "BEL.NS"),
        ("Mazagon Dock Shipbuilders", "MAZDOCK.NS"), ("Cochin Shipyard Ltd", "COCHINSHIP.NS"),
        ("Bharat Dynamics (BDL)", "BDL.NS"), ("Solar Industries India", "SOLARINDS.NS"),
        ("Data Patterns India", "DATAPATTNS.NS"), ("Paras Defence & Space", "PARAS.NS"),
        ("IRFC (Railway Finance)", "IRFC.NS"), ("RVNL (Rail Vikas Nigam)", "RVNL.NS"),
        ("IRCTC (Rail Catering)", "IRCTC.NS"), ("Titagarh Rail Systems", "TITAGARH.NS"),
        ("Jupiter Wagons Ltd", "JWL.NS"), ("RailTel Corporation", "RAILTEL.NS"),
        ("IRCON International", "IRCON.NS"), ("RITES Ltd", "RITES.NS")
    ],
    "🚗 EV, Automobile & Battery Tech (इलेक्ट्रिक वाहन व ऑटो)": [
        ("Tata Motors Ltd (EV Leader)", "TATAMOTORS.NS"), ("Mahindra & Mahindra (M&M)", "M&M.NS"),
        ("Maruti Suzuki India", "MARUTI.NS"), ("Bajaj Auto Ltd", "BAJAJ-AUTO.NS"),
        ("TVS Motor Company", "TVSMOTOR.NS"), ("Eicher Motors", "EICHERMOT.NS"),
        ("Hero MotoCorp", "HEROMOTOCO.NS"), ("Ola Electric Mobility", "OLAELEC.NS"),
        ("Olectra Greentech (EV Bus)", "OLECTRA.NS"), ("JBM Auto Ltd", "JBMA.NS"),
        ("Exide Industries (EV Battery)", "EXIDEIND.NS"), ("Amara Raja Energy", "ARE&M.NS"),
        ("Sona BLW Precision", "SONACOMS.NS"), ("Tesla Inc (US)", "TSLA"), ("Rivian Automotive (US)", "RIVN")
    ],
    "🌱 Green & Renewable Energy (सोलर, पवन व हरित ऊर्जा)": [
        ("Tata Power Company", "TATAPOWER.NS"), ("Suzlon Energy Ltd", "SUZLON.NS"),
        ("IREDA (Renewable Agency)", "IREDA.NS"), ("Adani Green Energy", "ADANIGREEN.NS"),
        ("Inox Wind Ltd", "INOXWIND.NS"), ("KPI Green Energy Ltd", "KPIGREEN.NS"),
        ("Waaree Energies Ltd", "WAAREE.NS"), ("Premier Energies Ltd", "PREMIERENE.NS"),
        ("NTPC Ltd / NTPC Green", "NTPC.NS"), ("Power Grid Corp", "POWERGRID.NS"),
        ("NHPC Ltd", "NHPC.NS"), ("SJVN Ltd", "SJVN.NS"), ("NextEra Energy (US)", "NEE"),
        ("First Solar Inc (US)", "FSLR")
    ],
    "🏦 PSU Banks & Govt Financials (सरकारी बैंक व वित्तीय संस्थान)": [
        ("Bank of India", "BANKINDIA.NS"), ("Bank of Baroda", "BANKBARODA.NS"),
        ("Canara Bank", "CANBK.NS"), ("State Bank of India (SBI)", "SBIN.NS"),
        ("Punjab National Bank (PNB)", "PNB.NS"), ("Union Bank of India", "UNIONBANK.NS"),
        ("Indian Bank", "INDIANB.NS"), ("Central Bank of India", "CENTRALBK.NS"),
        ("UCO Bank", "UCOBANK.NS"), ("Indian Overseas Bank", "IOB.NS"),
        ("Bank of Maharashtra", "MAHABANK.NS"), ("Punjab & Sind Bank", "PSB.NS"),
        ("IREDA", "IREDA.NS"), ("IRFC", "IRFC.NS"), ("PFC Limited", "PFC.NS"), ("REC Limited", "REC.NS")
    ],
    "🏢 NIFTY 50, NIFTY 200 & NIFTY 500": [
        ("Reliance Industries", "RELIANCE.NS"), ("TCS", "TCS.NS"), ("HDFC Bank", "HDFCBANK.NS"),
        ("ICICI Bank", "ICICIBANK.NS"), ("Infosys", "INFY.NS"), ("SBI", "SBIN.NS"),
        ("Bharti Airtel", "BHARTIARTL.NS"), ("ITC Ltd", "ITC.NS"), ("L&T", "LT.NS"),
        ("Sun Pharma", "SUNPHARMA.NS"), ("Titan Company", "TITAN.NS"), ("Bajaj Finance", "BAJFINANCE.NS"),
        ("UltraTech Cement", "ULTRACEMCO.NS"), ("Zomato Ltd", "ZOMATO.NS"), ("Jio Financial", "JIOFIN.NS")
    ],
    "💻 IT, Cloud & Artificial Intelligence (आईटी व AI)": [
        ("TCS", "TCS.NS"), ("Infosys", "INFY.NS"), ("HCL Tech", "HCLTECH.NS"),
        ("Wipro Ltd", "WIPRO.NS"), ("LTIMindtree", "LTIM.NS"), ("Tech Mahindra", "TECHM.NS"),
        ("Persistent Systems", "PERSISTENT.NS"), ("Coforge Ltd", "COFORGE.NS"),
        ("Tata Elxsi", "TATAELXSI.NS"), ("KPIT Tech", "KPITTECH.NS"), ("Microsoft Corp (US)", "MSFT"),
        ("Alphabet Google (US)", "GOOGL"), ("Amazon AWS (US)", "AMZN"), ("Meta Platforms (US)", "META"),
        ("Palantir Technologies (US)", "PLTR")
    ],
    "🇺🇸 US Mega Caps & Global Equities": [
        ("Apple Inc (US)", "AAPL"), ("Microsoft Corp (US)", "MSFT"), ("NVIDIA Corp (US)", "NVDA"),
        ("Alphabet Google (US)", "GOOGL"), ("Amazon.com (US)", "AMZN"), ("Meta Platforms (US)", "META"),
        ("Tesla Inc (US)", "TSLA"), ("Broadcom Inc (US)", "AVGO"), ("Palantir Tech (US)", "PLTR"),
        ("JPMorgan Chase (US)", "JPM"), ("Berkshire Hathaway (US)", "BRK-B")
    ]
}

# 4. Upcoming IPO Radar
UPCOMING_IPOS_DATA = [
    {"IPO Name": "Waaree Energies Limited", "Sector": "Solar / Renewable", "Price Band": "₹1,427 - ₹1,503", "Estimated GMP": "+95%", "Rating Review": "4.8/5 (Heavy Demand)", "AI Verdict": "🟢 STRONG APPLY"},
    {"IPO Name": "Hyundai Motor India", "Sector": "Automobile", "Price Band": "₹1,865 - ₹1,960", "Estimated GMP": "+8%", "Rating Review": "4.0/5 (Market Leader)", "AI Verdict": "🟢 APPLY FOR LONG TERM"},
    {"IPO Name": "Swiggy Limited", "Sector": "Quick Commerce", "Price Band": "₹371 - ₹390", "Estimated GMP": "+12%", "Rating Review": "3.8/5 (High Growth)", "AI Verdict": "🟡 APPLY FOR HIGH RISK"},
    {"IPO Name": "NTPC Green Energy Limited", "Sector": "PSU Renewable", "Price Band": "₹102 - ₹108", "Estimated GMP": "+25%", "Rating Review": "4.6/5 (Sovereign Backed)", "AI Verdict": "🟢 STRONG APPLY"}
]

# 5. Sidebar Settings
st.sidebar.markdown("### ⚙️ सेटिंग्स / Settings")
language = st.sidebar.radio("🌐 भाषा चुनें / Select Language:", ["Bilingual (हिंदी + English)", "हिंदी (Hindi)", "English"], index=0)
is_hindi = "हिंदी" in language
is_bilingual = "Bilingual" in language

def get_txt(hi, en):
    if is_bilingual: return f"{hi} | {en}"
    return hi if is_hindi else en

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 💰 {get_txt('पोर्टफोलियो व री-बाय (Re-Buy Averaging)', 'Portfolio & Re-Buy')}")
buy_price = st.sidebar.number_input(
    get_txt("आपका पुराना खरीद भाव (Your Old Buy Price):", "Your Buy Price:"),
    min_value=0.0, value=0.0, step=1.0,
    help="Yield on Cost और Re-Buy/Averaging Price निकालने के लिए दर्ज करें।"
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"### 🔐 {get_txt('एडमिन अनलॉक (Admin Access)', 'Admin Access')}")
admin_pass = st.sidebar.text_input(get_txt("एडमिन पासकोड दर्ज करें:", "Enter Admin Passcode:"), type="password")

ADMIN_PASSCODES = ["DEEPAK@1200", "ADMIN2026", "DEEPAK"]
is_admin = admin_pass.strip() in ADMIN_PASSCODES

if is_admin:
    st.sidebar.success("👑 एडमिन अनलॉक सक्रिय! (100% फ्री प्रीमियम एक्सेस)")
else:
    st.sidebar.info("ℹ️ प्रीमियम: ₹10/क्विक या ₹30/डिटेल्ड रिपोर्ट")

# 6. Top Banner & Title
st.markdown(
    """
    <div class="banner-ad">
        📢 SPONSORED / ADVERTISEMENT<br>
        ⚡ <b>Zero Brokerage Global & Indian Stock Investing</b> | <a href="#" target="_blank">Open Account Now</a>
    </div>
    """,
    unsafe_allow_html=True,
)

st.title("TradingView Pro | Global Stock & Fundamental AI Terminal")
st.caption("Semiconductor, EV, Railways, Defence, PSU Banks, NIFTY 50/500 & US Equities • TradingView Screener • Option Chain • AI Re-Buy Engine")

# 7. Sector & Universal Search
st.markdown(f"<div class='sec-header'>{get_txt('🔎 सेक्टर बास्केट व यूनिवर्सल सर्च', 'Sector Baskets & Universal Stock Search')}</div>", unsafe_allow_html=True)

col_sec, col_search = st.columns([1, 2])
with col_sec:
    selected_sector = st.selectbox(
        get_txt("📂 सेक्टर / इंडेक्स चुनें:", "Select Sector / Index:"),
        list(ALL_SECTORS_AND_INDICES.keys()), index=0
    )

sector_stocks = ALL_SECTORS_AND_INDICES[selected_sector]
sector_stock_map = {f"{name} [{ticker}]": ticker for name, ticker in sector_stocks}
sector_stock_options = list(sector_stock_map.keys()) + ["➕ Type Any Custom Symbol (अन्य कोई भी सिंबल लिखें)"]

with col_search:
    selected_stock_display = st.selectbox(
        get_txt("🔎 स्टॉक चुनें या नाम टाइप करें (Type to Search Any Stock):", "Select or Type Stock Name:"),
        options=sector_stock_options, index=0,
        help="Bank of India, Bank of Baroda, Canara Bank, Tata, Nvidia या कोई भी स्टॉक सर्च करें।"
    )

if selected_stock_display == "➕ Type Any Custom Symbol (अन्य कोई भी सिंबल लिखें)":
    custom_in = st.text_input(get_txt("सिंबल लिखें (उदा. BANKINDIA.NS, CANBK.NS, NVDA):", "Enter Symbol:"), value="BANKINDIA.NS").strip().upper()
    symbol = f"{custom_in}.NS" if ("." not in custom_in and not custom_in.startswith("^") and len(custom_in) > 3 and custom_in.isalpha()) else custom_in
else:
    symbol = sector_stock_map[selected_stock_display]

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

# Customization Expander
with st.expander(get_txt("🛠️ कस्टमाइज़ेशन विकल्प (Custom Columns & Sheets)", "Custom Columns & Sheets Settings")):
    cc1, cc2, cc3 = st.columns(3)
    inc_ohlc = cc1.checkbox(get_txt("OHLCV डेटा शीट शामिल करें", "Include OHLCV Sheet"), value=True)
    inc_div_sheet = cc2.checkbox(get_txt("डिविडेंड इतिहास शीट शामिल करें", "Include Dividend Sheet"), value=True)
    inc_summary = cc3.checkbox(get_txt("एग्जीक्यूटिव समरी शीट शामिल करें", "Include Executive Summary"), value=True)

# 8. Technical & Valuation Helper Functions
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

def evaluate_fundamental_health(info):
    score = 0
    factors = []
    op_margin = info.get("operatingMargins") or 0.0
    if op_margin > 0.15: score += 2; factors.append("✅ मजबूत ऑपरेटिंग मार्जिन (>15%)")
    elif op_margin > 0.08: score += 1; factors.append("⚖️ संतोषजनक ऑपरेटिंग मार्जिन")
    else: factors.append("⚠️ कम ऑपरेटिंग मार्जिन (<8%)")

    roe = info.get("returnOnEquity") or 0.0
    if roe > 0.15: score += 2; factors.append("✅ उत्कृष्ट ROE (>15%)")
    elif roe > 0.08: score += 1; factors.append("⚖️ मध्यम ROE")
    else: factors.append("⚠️ कमजोर ROE (<8%)")

    dte = info.get("debtToEquity")
    if dte is not None:
        if dte < 50: score += 2; factors.append("✅ सुरक्षित डेट-टू-इक्विटी (कम कर्ज)")
        elif dte < 120: score += 1; factors.append("⚖️ प्रबंधनीय कर्ज स्तर")
        else: factors.append("⚠️ उच्च कर्ज (High Debt/Equity)")
    else: score += 1

    fcf = info.get("freeCashflow") or 0
    if fcf > 0: score += 2; factors.append("✅ सकारात्मक फ्री कैश फ्लो (Positive FCF)")
    else: factors.append("⚠️ नकारात्मक / सीमित कैश फ्लो")

    pe = info.get("trailingPE") or 0
    if 0 < pe < 30: score += 2; factors.append("✅ उचित वैल्युएशन (P/E < 30)")
    elif pe >= 30: score += 1; factors.append("⚠️ उच्च प्रीमियम वैल्युएशन")

    health_pct = int((score / 10.0) * 100)
    if health_pct >= 70: category = "FUNDAMENTALLY VERY SOUND 🛡️ (अति मजबूत कंपनी)"; style_class = "fund-sound"
    elif health_pct >= 45: category = "FUNDAMENTALLY MODERATE ⚖️ (मध्यम / स्थिर कंपनी)"; style_class = "fund-mod"
    else: category = "FUNDAMENTALLY WEAK ⚠️ (कमजोर फंडामेंटल्स - सतर्क रहें)"; style_class = "fund-weak"

    return health_pct, category, style_class, factors

def fetch_option_chain_oi(ticker_obj, cmp):
    try:
        expirations = ticker_obj.options
        if not expirations: return None
        opt = ticker_obj.option_chain(expirations[0])
        calls, puts = opt.calls, opt.puts

        total_call_oi = calls["openInterest"].sum() if "openInterest" in calls else 0
        total_put_oi = puts["openInterest"].sum() if "openInterest" in puts else 0
        pcr = round(total_put_oi / (total_call_oi + 1e-9), 2)

        if pcr > 1.25: oi_sentiment = "BULLISH (मजबूत पुट राइटिंग / तेजी का रुख)"; oi_action = "🟢 BUY ON DIPS"
        elif pcr < 0.75: oi_sentiment = "BEARISH (मजबूत कॉल राइटिंग / दबाव का संकेत)"; oi_action = "🔴 SELL ON RISE"
        else: oi_sentiment = "NEUTRAL / RANGEBOUND (संतुलित दायरा)"; oi_action = "🟡 RANGE ACCUMULATION"

        max_call_strike = calls.loc[calls["openInterest"].idxmax()]["strike"] if not calls.empty and "openInterest" in calls else cmp * 1.05
        max_put_strike = puts.loc[puts["openInterest"].idxmax()]["strike"] if not puts.empty and "openInterest" in puts else cmp * 0.95
        option_fair_center = round((max_call_strike + max_put_strike) / 2, 2)

        return {
            "expiry": expirations[0], "total_call_oi": total_call_oi, "total_put_oi": total_put_oi,
            "pcr": pcr, "sentiment": oi_sentiment, "oi_action": oi_action,
            "call_resistance": max_call_strike, "put_support": max_put_strike,
            "option_fair_price": option_fair_center
        }
    except Exception:
        return None

# 9. TradingView Multi-Tab Screener
st.markdown(f"<div class='sec-header'>{get_txt('📊 TradingView लाइव स्टॉक स्क्रीनर ग्रिड', 'TradingView Live Stock Screener Grid')}</div>", unsafe_allow_html=True)

screener_tabs = st.tabs(["Overview", "Technicals", "Valuation", "Dividends & Margins", "🚀 Upcoming IPO Radar"])

with screener_tabs[0]:
    if st.button("⚡ रन ओवरव्यू स्क्रीनर स्कैन (Run Overview Scan)", key="run_overview_scan"):
        with st.spinner(f"Scanning stocks in {selected_sector}..."):
            rows = []
            for s_name, s_ticker in sector_stocks[:20]:
                try:
                    s_t = yf.Ticker(s_ticker)
                    s_h = s_t.history(period="3mo")
                    s_inf = s_t.info or {}
                    if not s_h.empty:
                        c_p = round(float(s_h["Close"].iloc[-1]), 2)
                        p_c = round(float(s_h["Close"].iloc[-2]), 2) if len(s_h) > 1 else c_p
                        chg = round(((c_p - p_c) / p_c) * 100, 2)
                        vol = f"{s_h['Volume'].iloc[-1]/1e6:.2f}M" if "Volume" in s_h else "N/A"
                        mkt_c = f"₹{s_inf.get('marketCap', 0)/1e12:.2f}T" if s_inf.get('marketCap') else "N/A"
                        p_e = round(s_inf.get('trailingPE', 0), 2) if s_inf.get('trailingPE') else "N/A"
                        eps_val = round(s_inf.get('trailingEps', 0), 2) if s_inf.get('trailingEps') else "N/A"
                        div_y = f"{s_inf.get('dividendYield', 0)*100:.2f}%" if s_inf.get('dividendYield') else "0.00%"
                        rating = str(s_inf.get('recommendationKey', 'Buy')).replace('_', ' ').title()
                        
                        rows.append({
                            "Symbol": s_ticker, "Company Name": s_name, "Price": c_p, "Chg %": f"{chg:+}%",
                            "Vol": vol, "Mkt Cap": mkt_c, "P/E": p_e, "EPS (TTM)": eps_val,
                            "Div Yield %": div_y, "Sector": s_inf.get('sector', selected_sector),
                            "Analyst Rating": f"⭐ {rating}"
                        })
                except Exception:
                    continue
            if rows: st.dataframe(pd.DataFrame(rows), use_container_width=True)

with screener_tabs[1]:
    if st.button("⚡ रन टेक्निकल इंडिकेटर स्कैन (Run Technicals Scan)", key="run_tech_scan"):
        with st.spinner("Calculating RSI, MACD & Signals..."):
            tech_rows = []
            for s_name, s_ticker in sector_stocks[:20]:
                try:
                    s_t = yf.Ticker(s_ticker)
                    s_h = s_t.history(period="3mo")
                    if not s_h.empty and len(s_h) >= 15:
                        c_p = round(float(s_h["Close"].iloc[-1]), 2)
                        r_val = round(float(calculate_rsi(s_h["Close"]).iloc[-1]), 1)
                        m_line, s_line = calculate_macd(s_h["Close"])
                        m_val = round(float(m_line.iloc[-1]), 2)
                        
                        if r_val <= 35: act = "🟢 Strong Buy (Oversold)"
                        elif r_val >= 70: act = "🔴 Strong Sell (Overbought)"
                        elif m_val > 0 and r_val > 55: act = "🟢 Buy (Momentum)"
                        else: act = "🟡 Neutral / Hold"

                        tech_rows.append({
                            "Symbol": s_ticker, "Company Name": s_name, "Price": c_p,
                            "RSI (14)": r_val, "MACD": m_val, "AI Action Signal": act
                        })
                except Exception:
                    continue
            if tech_rows: st.dataframe(pd.DataFrame(tech_rows).sort_values(by="RSI (14)", ascending=False), use_container_width=True)

with screener_tabs[2]:
    if st.button("⚡ रन वैल्युएशन स्कैन (Run Valuation Scan)", key="run_val_scan"):
        with st.spinner("Fetching P/E, P/B & Intrinsic Values..."):
            val_rows = []
            for s_name, s_ticker in sector_stocks[:20]:
                try:
                    s_t = yf.Ticker(s_ticker)
                    s_inf = s_t.info or {}
                    c_p = s_inf.get('currentPrice') or s_inf.get('regularMarketPrice') or 0.0
                    eps_v = s_inf.get('trailingEps') or 0.0
                    bv_v = s_inf.get('bookValue') or 0.0
                    iv_v = calculate_intrinsic_value(eps_v, bv_v) if eps_v and bv_v else "N/A"

                    val_rows.append({
                        "Symbol": s_ticker, "Company Name": s_name, "CMP": c_p,
                        "P/E": round(s_inf.get('trailingPE', 0), 2) if s_inf.get('trailingPE') else "N/A",
                        "P/B": round(s_inf.get('priceToBook', 0), 2) if s_inf.get('priceToBook') else "N/A",
                        "Book Value": bv_v, "Intrinsic Value (Fair)": iv_v
                    })
                except Exception:
                    continue
            if val_rows: st.dataframe(pd.DataFrame(val_rows), use_container_width=True)

with screener_tabs[3]:
    if st.button("⚡ रन डिविडेंड व मार्जिन स्कैन (Run Margins Scan)", key="run_div_scan"):
        with st.spinner("Fetching Margins and Dividend Yields..."):
            div_rows = []
            for s_name, s_ticker in sector_stocks[:20]:
                try:
                    s_inf = yf.Ticker(s_ticker).info or {}
                    div_rows.append({
                        "Symbol": s_ticker, "Company Name": s_name,
                        "Div Yield": f"{s_inf.get('dividendYield', 0)*100:.2f}%" if s_inf.get('dividendYield') else "0.00%",
                        "Operating Margin": f"{s_inf.get('operatingMargins', 0)*100:.2f}%" if s_inf.get('operatingMargins') else "N/A",
                        "ROE": f"{s_inf.get('returnOnEquity', 0)*100:.2f}%" if s_inf.get('returnOnEquity') else "N/A"
                    })
                except Exception:
                    continue
            if div_rows: st.dataframe(pd.DataFrame(div_rows), use_container_width=True)

with screener_tabs[4]:
    st.dataframe(pd.DataFrame(UPCOMING_IPOS_DATA), use_container_width=True)

# 10. Fetch Single Stock Data Payload
def fetch_stock_payload(ticker_symbol, period_val, s_date, e_date):
    try:
        t = yf.Ticker(ticker_symbol)
        h = t.history(period=period_val) if period_val else t.history(start=s_date, end=e_date)
        if h.empty: h = t.history(period="1y")
        max_h = t.history(period="max")
        info = t.info or {}
        divs = t.dividends if hasattr(t, "dividends") else pd.Series(dtype=float)
        ath = max_h["High"].max() if not max_h.empty else (h["High"].max() if not h.empty else None)
        return t, h, info, ath, divs
    except Exception:
        return None, None, None, None, None

def generate_premium_excel(summary_data, hist_df, div_df, inc_s, inc_o, inc_d):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if inc_s and summary_data: pd.DataFrame(summary_data).to_excel(writer, sheet_name="Executive Summary", index=False)
        if inc_o and not hist_df.empty:
            h_exp = hist_df.reset_index()
            if "Date" in h_exp.columns: h_exp["Date"] = h_exp["Date"].dt.strftime("%Y-%m-%d")
            h_exp.to_excel(writer, sheet_name="Historical OHLC Data", index=False)
        if inc_d and not div_df.empty:
            d_exp = div_df.reset_index()
            if "Date" in d_exp.columns: d_exp["Date"] = d_exp["Date"].dt.strftime("%Y-%m-%d")
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

# 11. Execution & Single Stock Analytics
if symbol:
    with st.spinner(f"Fetching Live Market Analytics for {symbol}..."):
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
        sector = stock_info.get("sector", selected_sector)
        industry = stock_info.get("industry", "Equity Market")

        fund_score, fund_verdict, fund_class, fund_factors = evaluate_fundamental_health(stock_info)

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

        rsi_sig = "OVERSOLD (BUY)" if latest_rsi < 35 else ("OVERBOUGHT (SELL)" if latest_rsi > 70 else "NEUTRAL")
        macd_sig = "BULLISH CROSSOVER (BUY)" if latest_macd > latest_sig else "BEARISH CROSSOVER (SELL)"
        trend_sig = "BULLISH (Above 50 SMA)" if cmp_price > sma_50_val else "BEARISH (Below 50 SMA)"

        oi_data = fetch_option_chain_oi(ticker_obj, cmp_price)

        score = 0
        if latest_rsi < 45: score += 1.5
        elif latest_rsi < 60: score += 1.0
        if latest_macd > latest_sig: score += 1.5
        if cmp_price > sma_50_val: score += 1.0
        if down_from_52w < -15: score += 1.0
        if fund_score >= 60: score += 1.5
        if oi_data and "BULLISH" in oi_data["sentiment"]: score += 1.0

        win_prob = round(min(max((score / 7.5) * 100, 25.0), 93.0), 1)
        ai_action = "STRONG BUY 🚀 (जोरदार खरीदारी)" if win_prob >= 75 else ("BUY 📈 (खरीदें)" if win_prob >= 58 else ("HOLD ⚖️ (बनाए रखें)" if win_prob >= 45 else "SELL / AVOID 📉 (बेचें / बचें)"))

        intrinsic_val = calculate_intrinsic_value(eps, book_val) if eps and book_val else None
        if intrinsic_val:
            ai_fair_buy_price = round((intrinsic_val * 0.85 + cmp_price * 0.95) / 2, 2)
            ai_max_buy_price = round(intrinsic_val * 0.95, 2)
        else:
            ai_fair_buy_price = round(cmp_price * 0.95, 2)
            ai_max_buy_price = round(cmp_price * 0.98, 2)

        if buy_price > 0:
            if cmp_price < buy_price:
                suggested_rebuy_price = round(cmp_price * 0.98, 2)
                rebuy_advice = f"🟢 स्टॉक आपके खरीद भाव से {((buy_price-cmp_price)/buy_price*100):.1f}% नीचे है। `{currency} {suggested_rebuy_price}` पर एक्युमुलेट/एवरेज करें।"
            else:
                suggested_rebuy_price = round(max(cmp_price * 0.96, buy_price), 2)
                rebuy_advice = f"⚖️ स्टॉक आपके खरीद भाव से मुनाफे में है। पिरामिडिंग हेतु डिप पर `{currency} {suggested_rebuy_price}` पर जोड़ें।"
        else:
            suggested_rebuy_price = ai_fair_buy_price
            rebuy_advice = "Sidebar में अपना पुराना खरीद भाव दर्ज करके री-बाय स्तर देखें।"

        entry_lvl = round(cmp_price * 0.985, 2)
        sl_lvl = round(cmp_price * 0.94, 2)
        tgt_1 = round(cmp_price * 1.08, 2)
        tgt_2 = round(cmp_price * 1.15, 2)

        analyst_recom = str(stock_info.get("recommendationKey", "BUY")).upper()
        target_mean = stock_info.get("targetMeanPrice", round(cmp_price * 1.14, 2))

        total_lifetime_div = float(df_div.sum()) if not df_div.empty else 0.0
        yield_on_cost = (div_rate / buy_price * 100) if buy_price > 0 else None

        # Stock Header
        st.markdown("---")
        st.subheader(f"🏢 {long_name} ({symbol})")
        st.caption(f"Sector: **{sector}** | Industry: **{industry}** | Currency: **{currency}**")

        # 0. Fundamental Soundness
        st.markdown(f"<div class='sec-header'>{get_txt('🛡️ AI फंडामेंटल हेल्थ व कंपनी साउंडनेस (Fundamental Soundness & Health)', 'AI Fundamental Soundness & Health Score')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='{fund_class}'>📊 <b>कंपनी स्थिति:</b> {fund_verdict} | <b>AI हेल्थ स्कोर:</b> {fund_score}/100</div>", unsafe_allow_html=True)

        fcol1, fcol2, fcol3 = st.columns(3)
        fcol1.info(f"💎 **AI सही खरीद मूल्य (Fair Buy Price):** `{currency} {ai_fair_buy_price:,.2f}`\n\n*(इस स्तर पर रिस्क न्यूनतम है)*")
        fcol2.info(f"🛑 **अधिकतम खरीद सीमा (Max Buy Limit):** `{currency} {ai_max_buy_price:,.2f}`\n\n*(इसके ऊपर ओवरप्राइस्ड माना जाएगा)*")
        factors_txt = "\n".join(fund_factors[:3])
        fcol3.success(f"📋 **मुख्य फंडामेंटल कारक:**\n\n{factors_txt}")

        # 0.1 Re-Buy Price
        st.markdown(f"<div class='sec-header'>{get_txt('🎯 AI री-बाय / एवरेजिंग कैलकुलेटर (Re-Buy Price for Existing Holders)', 'AI Re-Buy & Position Averaging Price')}</div>", unsafe_allow_html=True)
        rb_col1, rb_col2 = st.columns([1, 2])
        rb_col1.metric(label=get_txt("सुझाया गया री-बाय स्तर (Re-Buy Price)", "Re-Buy Price"), value=f"{currency} {suggested_rebuy_price:,.2f}")
        rb_col2.write(f"**सलाह व स्थिति:** {rebuy_advice}")

        # 1. Price Action & ATH
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

        # 2. Valuation
        st.markdown(f"<div class='sec-header'>{get_txt('वैल्युएशन एवं फंडामेंटल्स (P/E & P/B Multiples)', 'Valuation & Fundamentals')}</div>", unsafe_allow_html=True)
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Company P/E", f"{company_pe:.2f}" if company_pe else "N/A")
        v2.metric("Industry P/E", str(industry_pe))
        v3.metric("P/B Ratio", f"{pb_ratio:.2f}" if pb_ratio else "N/A")
        v4.metric("EPS (TTM)", f"{currency} {eps:.2f}" if eps else "N/A")
        v5.metric("Dividend Yield (CMP)", f"{div_yield:.2f}%")

        # 3. Dividend Intelligence & Custom Date Filter
        st.markdown(f"<div class='sec-header'>{get_txt('💰 डिविडेंड विश्लेषण एवं पूंजी यील्ड (Dividend Analytics & Yield on Cost)', 'Dividend Analytics & Yield on Cost')}</div>", unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.metric(get_txt("लाइफटाइम कुल डिविडेंड", "Lifetime Total Div"), f"{currency} {total_lifetime_div:,.2f}")
        d2.metric(get_txt("वार्षिक डिविडेंड दर (TTM)", "Annual Div Rate (TTM)"), f"{currency} {div_rate:,.2f}")
        d3.metric(get_txt("बुक वैल्यू (Book Value)", "Book Value"), f"{currency} {book_val:,.2f}" if book_val else "N/A")
        if yield_on_cost is not None:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड (Yield on Cost)", "Yield on Cost (Your Buy)"), f"{yield_on_cost:.2f}%", f"Buy: {currency} {buy_price}")
        else:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड", "Yield on Cost"), "Sidebar में दर्ज करें")

        # Dividend Date Filter Expander
        with st.expander(get_txt("📅 कस्टम तारीख अनुसार डिविडेंड कैलकुलेटर", "Custom Date Range Dividend Calculator")):
            div_c1, div_c2 = st.columns(2)
            c_start = div_c1.date_input("Dividend Filter Start", value=datetime.date(2020, 1, 1), key="div_c_start")
            c_end = div_c2.date_input("Dividend Filter End", value=datetime.date.today(), key="div_c_end")
            if not df_div.empty:
                div_clean = df_div.copy()
                div_clean.index = div_clean.index.tz_localize(None)
                mask = (div_clean.index.date >= c_start) & (div_clean.index.date <= c_end)
                range_div_sum = div_clean[mask].sum()
                st.write(f"**{c_start} से {c_end} के बीच कुल डिविडेंड:** `{currency} {range_div_sum:,.2f}`")
                st.dataframe(div_clean[mask], use_container_width=True)

        # 4. TradingView Multi-Panel Chart
        st.markdown(f"<div class='sec-header'>📈 TradingView Pro Technical Chart (SMA, BB, MACD & RSI)</div>", unsafe_allow_html=True)
        fig = make_subplots(
            rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.03,
            subplot_titles=("Price & Bollinger Bands", "MACD (12, 26, 9)", "RSI (14)"),
            row_heights=[0.6, 0.2, 0.2]
        )

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

        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["MACD"], name="MACD Line", line=dict(color="#2962ff", width=1.5)), row=2, col=1)
        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["MACD_Sig"], name="Signal Line", line=dict(color="#ff3d60", width=1.5)), row=2, col=1)

        fig.add_trace(go.Scatter(x=df_hist.index, y=df_hist["RSI"], name="RSI", line=dict(color="#9b59b6", width=1.5)), row=3, col=1)
        fig.add_hline(y=70, line_dash="dash", line_color="red", row=3, col=1)
        fig.add_hline(y=30, line_dash="dash", line_color="green", row=3, col=1)

        fig.update_layout(height=650, xaxis_rangeslider_visible=False, template="plotly_white", margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True)

        with st.expander(get_txt("📋 पूर्ण डेटा तालिका देखें (View Full OHLC Table)", "View Full OHLC Table")):
            st.dataframe(df_hist, use_container_width=True)

        # 5. Paywall & Premium AI Report
        st.markdown(f"<div class='sec-header'>{get_txt('💎 AI एक्सपर्ट रिपोर्ट व F&O / OI विश्लेषण (Buy/Sell Recommendation Engine)', 'AI Expert Report & F&O Analytics')}</div>", unsafe_allow_html=True)

        if "unlocked_quick" not in st.session_state: st.session_state.unlocked_quick = False
        if "unlocked_detailed" not in st.session_state: st.session_state.unlocked_detailed = False

        has_access_quick = is_admin or st.session_state.unlocked_quick
        has_access_detailed = is_admin or st.session_state.unlocked_detailed

        if not (has_access_quick or has_access_detailed):
            st.markdown(
                """
                <div class="premium-box">
                    <h3>🔒 प्रीमियम AI खरीद/बिक्री सलाह, F&O/OI डेरिवेटिव्स व संस्थागत रिपोर्ट लॉक है</h3>
                    <p>यह रिपोर्ट टॉप इंडिकेटर्स (RSI, MACD, Bollinger Bands), AI प्रोबेबिलिटी स्कोर, <b>Actionable Buy/Sell Verdict</b>, टार्गेट, स्टॉप-लॉस, सही खरीद मूल्य, ऑप्शन चेन फेयर वैल्यू व ब्रोकरेज रेटिंग्स का लाइव विश्लेषण करती है।</p>
                    <ul>
                        <li><b>₹10 / Quick Report:</b> AI वर्डिक्ट + टेक्निकल इंडिकेटर सारांश</li>
                        <li><b>₹30 / Detailed Analysis:</b> पूर्ण AI प्रोबेबिलिटी + Re-Buy/Averaging + F&O/OI डेरिवेटिव्स + ब्रोकरेज रेटिंग्स + प्रीमियम एक्सेल एक्सपोर्ट</li>
                        <li>👑 <b>एडमिन:</b> साइडबार में पासकोड डालकर 100% फ्री अनलॉक करें।</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )
            p1, p2 = st.columns(2)
            if p1.button("💳 ₹10 में Quick AI Report अनलॉक करें"):
                st.session_state.unlocked_quick = True
                st.rerun()
            if p2.button("👑 ₹30 में Full Detailed Analysis अनलॉक करें"):
                st.session_state.unlocked_detailed = True
                st.rerun()

        if has_access_quick or has_access_detailed:
            st.success("✅ प्रीमियम AI खरीद/बिक्री व F&O रिपोर्ट अनलॉक!")

            st.markdown("#### 🎯 AI खरीद/बिक्री फैसला व ब्रोकरेज कंसेंसस")
            r1, r2, r3 = st.columns(3)
            r1.metric("📌 AI फैसला (Action)", ai_action)
            r2.metric("📊 विन प्रोबेबिलिटी स्कोर", f"{win_prob}%", "फंडामेंटल, F&O व टेक्निकल डेटा पर")
            r3.metric("🏢 ब्रोकरेज रेटिंग कंसेंसस", analyst_recom, f"Target: {currency} {target_mean:,.1f}")

            st.markdown("#### 📍 AI सुझाई गई कीमतें व लेवल्स (AI Calculated Levels)")
            l1, l2, l3, l4 = st.columns(4)
            l1.metric("📥 सही खरीद मूल्य (Fair Buy)", f"{currency} {ai_fair_buy_price:,.2f}")
            l2.metric("🛑 स्टॉप-लॉस (Stop-Loss)", f"{currency} {sl_lvl:,.2f}", "-6% Buffer", delta_color="inverse")
            l3.metric("🎯 टार्गेट 1", f"{currency} {tgt_1:,.2f}", "+8% Target")
            l4.metric("🚀 टार्गेट 2", f"{currency} {tgt_2:,.2f}", "+15% Target")

            if oi_data:
                st.markdown("#### 📈 Live Option Chain, Open Interest (OI) व डेरिवेटिव्स विश्लेषण")
                o1, o2, o3, o4 = st.columns(4)
                o1.metric("Put-Call Ratio (PCR)", f"{oi_data['pcr']}")
                o2.metric("OI Action Signal", oi_data["oi_action"])
                o3.metric("Option Fair Price Center", f"{currency} {oi_data['option_fair_price']}")
                o4.metric("Support / Resistance", f"{oi_data['put_support']} / {oi_data['call_resistance']}")
                st.info(f"💡 **F&O / OI सेंटीमेंट इंटरप्रिटेशन:** {oi_data['sentiment']} (Expiry: {oi_data['expiry']})")

            st.markdown("#### ⚙️ तकनीकी इंडिकेटर्स सिग्नल (Technical Signals)")
            t1, t2, t3 = st.columns(3)
            t1.metric("RSI (14-Day)", f"{latest_rsi:.1f}", rsi_sig)
            t2.metric("MACD Status", f"{latest_macd:.2f}", macd_sig)
            t3.metric("Overall Trend", trend_sig)

        # Excel Summary Rows
        summary_rows = [
            {"Field": "--- GENERAL OVERVIEW ---", "Value": ""},
            {"Field": "Company Name", "Value": str(long_name)},
            {"Field": "Symbol", "Value": str(symbol)},
            {"Field": "Sector / Industry", "Value": f"{sector} / {industry}"},
            {"Field": "--- AI FUNDAMENTAL HEALTH ---", "Value": ""},
            {"Field": "Fundamental Health Status", "Value": fund_verdict},
            {"Field": "Fundamental Health Score", "Value": f"{fund_score}/100"},
            {"Field": "AI Fair Buy Price", "Value": f"{currency} {ai_fair_buy_price:,.2f}"},
            {"Field": "AI Max Buy Limit", "Value": f"{currency} {ai_max_buy_price:,.2f}"},
            {"Field": "AI Re-Buy / Averaging Price", "Value": f"{currency} {suggested_rebuy_price:,.2f}"},
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
            {"Field": "--- F&O / OPTION CHAIN INSIGHTS ---", "Value": ""},
            {"Field": "PCR (Put-Call Ratio)", "Value": str(oi_data["pcr"]) if oi_data else "N/A"},
            {"Field": "OI Sentiment", "Value": str(oi_data["sentiment"]) if oi_data else "N/A"},
            {"Field": "Option Fair Price", "Value": f"{currency} {oi_data['option_fair_price']}" if oi_data else "N/A"},
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

        excel_data = generate_premium_excel(summary_rows, df_hist, df_div, inc_summary, inc_ohlc, inc_div_sheet)

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
