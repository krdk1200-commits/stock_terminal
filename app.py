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
    page_title="Global Stock & Fundamental Terminal",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling (Retaining Admin/Banner + Metric Cards)
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
    </style>
    """,
    unsafe_allow_html=True,
)

# --- THEMATIC & SUGGESTION DATABASE ---
THEMATIC_STOCK_DATA = {
    "🔥 All Popular Stocks & Indices (लोकप्रिय स्टॉक्स)": [
        ("NIFTY 50 Index (India)", "^NSEI"),
        ("SENSEX Index (India)", "^BSESN"),
        ("BANK NIFTY (India)", "^NSEBANK"),
        ("Reliance Industries (RIL)", "RELIANCE.NS"),
        ("Tata Consultancy Services (TCS)", "TCS.NS"),
        ("Tata Motors Ltd", "TATAMOTORS.NS"),
        ("Tata Steel Ltd", "TATASTEEL.NS"),
        ("Tata Power Co Ltd", "TATAPOWER.NS"),
        ("Tata Technologies Ltd", "TATATECH.NS"),
        ("Tata Elxsi Ltd", "TATAELXSI.NS"),
        ("Tata Consumer Products", "TATACONSUM.NS"),
        ("HDFC Bank Ltd", "HDFCBANK.NS"),
        ("ICICI Bank Ltd", "ICICIBANK.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"),
        ("Infosys Ltd", "INFY.NS"),
        ("ITC Ltd", "ITC.NS"),
        ("Larsen & Toubro (L&T)", "LT.NS"),
        ("Apple Inc. (US)", "AAPL"),
        ("Microsoft Corp (US)", "MSFT"),
        ("NVIDIA Corp (US)", "NVDA"),
        ("Tesla Inc. (US)", "TSLA"),
        ("Alphabet Google (US)", "GOOGL"),
        ("Amazon.com Inc (US)", "AMZN"),
        ("Meta Platforms (Facebook)", "META"),
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
        ("Tesla Inc. (Global EV - US)", "TSLA"),
        ("Rivian Automotive (US)", "RIVN"),
        ("BYD Company (ADR - US)", "BYDDY"),
        ("Lucid Group (US)", "LCID"),
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
    ],
    "🛡️ Defence & Aerospace (रक्षा एवं अंतरिक्ष)": [
        ("Hindustan Aeronautics (HAL)", "HAL.NS"),
        ("Bharat Electronics (BEL)", "BEL.NS"),
        ("Mazagon Dock Shipbuilders", "MAZDOCK.NS"),
        ("Cochin Shipyard Ltd", "COCHINSHIP.NS"),
        ("Bharat Dynamics Ltd (BDL)", "BDL.NS"),
        ("Solar Industries India", "SOLARINDS.NS"),
        ("Data Patterns India Ltd", "DATAPATTNS.NS"),
        ("Lockheed Martin (US Defence)", "LMT"),
        ("RTX Raytheon Technologies (US)", "RTX"),
        ("Boeing Company (US)", "BA"),
    ],
    "🚆 Railways & Infra (रेलवे व इंफ्रास्ट्रक्चर)": [
        ("IRFC (Indian Railway Finance)", "IRFC.NS"),
        ("RVNL (Rail Vikas Nigam)", "RVNL.NS"),
        ("IRCTC (Rail Catering & Tourism)", "IRCTC.NS"),
        ("Titagarh Rail Systems", "TITAGARH.NS"),
        ("Jupiter Wagons Ltd", "JWL.NS"),
        ("IRCON International Ltd", "IRCON.NS"),
        ("RailTel Corporation", "RAILTEL.NS"),
        ("Larsen & Toubro (L&T)", "LT.NS"),
    ],
    "💻 IT, Cloud & Artificial Intelligence (आईटी व AI)": [
        ("Tata Consultancy Services (TCS)", "TCS.NS"),
        ("Infosys Ltd", "INFY.NS"),
        ("HCL Technologies Ltd", "HCLTECH.NS"),
        ("Wipro Ltd", "WIPRO.NS"),
        ("LTIMindtree Ltd", "LTIM.NS"),
        ("Persistent Systems Ltd", "PERSISTENT.NS"),
        ("KPIT Technologies (Auto Tech)", "KPITTECH.NS"),
        ("Microsoft Corp (US)", "MSFT"),
        ("Alphabet Inc (Google - US)", "GOOGL"),
        ("Amazon Web Services (AWS - US)", "AMZN"),
        ("Meta Platforms (US)", "META"),
        ("Palantir Technologies (US)", "PLTR"),
        ("Salesforce Inc. (US)", "CRM"),
    ],
    "🏦 Banking, PSU & Financial Services (बैंकिंग एवं वित्त)": [
        ("HDFC Bank Ltd", "HDFCBANK.NS"),
        ("State Bank of India (SBI)", "SBIN.NS"),
        ("ICICI Bank Ltd", "ICICIBANK.NS"),
        ("Kotak Mahindra Bank", "KOTAKBANK.NS"),
        ("Axis Bank Ltd", "AXISBANK.NS"),
        ("Bank of Baroda", "BANKBARODA.NS"),
        ("Bajaj Finance Ltd", "BAJFINANCE.NS"),
        ("JPMorgan Chase & Co (US)", "JPM"),
        ("Bank of America (US)", "BAC"),
    ],
}

# --- SIDEBAR (Original Setup + Preserved Controls) ---
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
st.sidebar.markdown(f"### 🌍 {get_txt('लोकप्रिय इंडेक्स', 'Global Indices')}")
index_choice = st.sidebar.selectbox(
    get_txt("इंडेक्स चुनें / Select Index:", "Select Index:"),
    ["-- Manual / सिंबल दर्ज करें --", "NIFTY 50", "SENSEX", "BANK NIFTY", "NASDAQ 100", "S&P 500"],
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

# --- TOP BANNER (Preserved Ad/Sponsored Banner) ---
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
st.title(get_txt("Global Stock Terminal | वैश्विक मार्केट टर्मिनल", "Global Stock Terminal"))
st.caption(get_txt("Indian & US Stocks / Indices Analysis | भारतीय एवं अमेरिकी स्टॉक्स विश्लेषण", "Indian & US Stocks / Indices Analysis"))

# --- SEARCH & INPUT CONTROLS ---
scol1, scol2 = st.columns([1, 2])
with scol1:
    selected_theme = st.selectbox(
        get_txt("📂 सेक्टर / थीम बास्केट चुनें:", "Select Sector / Theme Basket:"),
        list(THEMATIC_STOCK_DATA.keys()),
    )

stock_list_for_theme = THEMATIC_STOCK_DATA[selected_theme]
stock_display_map = {f"{name} [{ticker}]": ticker for name, ticker in stock_list_for_theme}
stock_options = list(stock_display_map.keys()) + ["➕ Manual Custom Symbol (अन्य सिंबल लिखें)"]

with scol2:
    selected_stock_display = st.selectbox(
        get_txt("🔎 कंपनी / स्टॉक का नाम टाइप करके खोजें (Type to Search):", "Type Company / Stock Name to Search:"),
        options=stock_options,
        index=0,
        help="यहाँ नाम टाइप करें (उदा. Tata, Microsoft, M, A, Reliance) - नीचे तुरंत लाइव सजेशन ड्रॉपडाउन आएगा।",
    )

# Determine final symbol
if index_choice != "-- Manual / सिंबल दर्ज करें --":
    idx_map = {
        "NIFTY 50": "^NSEI",
        "SENSEX": "^BSESN",
        "BANK NIFTY": "^NSEBANK",
        "NASDAQ 100": "^NDX",
        "S&P 500": "^GSPC",
    }
    symbol = idx_map.get(index_choice, "RELIANCE.NS")
elif selected_stock_display == "➕ Manual Custom Symbol (अन्य सिंबल लिखें)":
    symbol = st.text_input(
        get_txt("Stock / Index Symbol (भारतीय या अमेरिकी सिंबल):", "Stock / Index Symbol:"),
        value="TATAMOTORS.NS",
    ).strip().upper()
else:
    symbol = stock_display_map[selected_stock_display]

# Duration / Range Options (Preserved Standard & Custom Range)
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

# Customization Expander (Preserved)
with st.expander(get_txt("🛠️ कस्टमाइज़ेशन विकल्प (Custom Columns & Sheets)", "Custom Columns & Sheets Settings")):
    cc1, cc2, cc3 = st.columns(3)
    inc_ohlc = cc1.checkbox(get_txt("OHLCV डेटा शीट शामिल करें", "Include OHLCV Sheet"), value=True)
    inc_div_sheet = cc2.checkbox(get_txt("डिविडेंड इतिहास शीट शामिल करें", "Include Dividend Sheet"), value=True)
    inc_summary = cc3.checkbox(get_txt("एग्जीक्यूटिव समरी शीट शामिल करें", "Include Executive Summary"), value=True)

# Helper Functions
def calculate_intrinsic_value(eps, book_value):
    try:
        if eps > 0 and book_value > 0:
            return round(np.sqrt(22.5 * eps * book_value), 2)
    except Exception:
        pass
    return None

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

# Execution Button / Flow
analyze_clicked = st.button(get_txt("🚀 Analyze / डेटा निकालें", "🚀 Analyze / Fetch Data"), use_container_width=True)

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

        # Summary for Excel
        summary_rows = [
            {"Field": "--- GENERAL OVERVIEW ---", "Value": ""},
            {"Field": "Company Name", "Value": str(long_name)},
            {"Field": "Symbol", "Value": str(symbol)},
            {"Field": "Sector / Industry", "Value": f"{sector} / {industry}"},
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
            file_name=f"{symbol}_Executive_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error(get_txt("डेटा प्राप्त करने में असमर्थ। कृपया सिंबल की जाँच करें।", "Unable to fetch data. Please check symbol."))
