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

# Custom Styling
st.markdown(
    """
    <style>
    .sec-header {
        font-size: 1.1rem;
        font-weight: bold;
        padding-bottom: 5px;
        margin-top: 20px;
        margin-bottom: 10px;
        border-bottom: 2px solid #1B365D;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- SIDEBAR ---
st.sidebar.title("⚙️ सेट्स / Settings")
language = st.sidebar.radio(
    "🌐 भाषा चुनें / Language:", ["Bilingual (हिंदी + English)", "English", "हिंदी (Hindi)"]
)

is_hindi = "हिंदी" in language
is_bilingual = "Bilingual" in language

def get_txt(hi, en):
    if is_bilingual:
        return f"{hi} | {en}"
    return hi if is_hindi else en

st.sidebar.markdown("---")
st.sidebar.subheader(get_txt("पोर्टफोलियो इनपुट (Portfolio Input)", "Portfolio Input"))
buy_price = st.sidebar.number_input(
    get_txt("आपका खरीद भाव (Your Buy Price):", "Your Buy Price:"),
    min_value=0.0,
    value=0.0,
    step=1.0,
    help="Yield on Cost निकालने के लिए अपना खरीद मूल्य दर्ज करें।"
)

# --- MAIN PAGE INPUTS ---
st.title(get_txt("🌍 वैश्विक मार्केट व फंडामेंटल टर्मिनल", "Global Stock & Fundamental Terminal"))

col1, col2 = st.columns([2, 1])
with col1:
    symbol = st.text_input(
        get_txt("स्टॉक / इंडेक्स सिंबल दर्ज करें:", "Enter Stock / Index Symbol:"),
        value="RELIANCE.NS",
    ).strip().upper()

with col2:
    period = st.selectbox(
        get_txt("चार्ट अवधि / Chart Duration:", "Select Historical Duration:"),
        ["1mo", "3mo", "6mo", "1y", "2y", "5y", "max"],
        index=3,
    )

# --- HELPER FUNCTIONS ---
def calculate_intrinsic_value(eps, book_value):
    try:
        if eps > 0 and book_value > 0:
            return round(np.sqrt(22.5 * eps * book_value), 2)
    except Exception:
        pass
    return None

def fetch_data(ticker_symbol, duration):
    try:
        ticker = yf.Ticker(ticker_symbol)
        hist = ticker.history(period=duration)
        max_hist = ticker.history(period="max")
        info = ticker.info
        dividends = ticker.dividends
        if hist.empty:
            return None, None, None, None, None
        ath = max_hist["High"].max() if not max_hist.empty else None
        return ticker, hist, info, ath, dividends
    except Exception:
        return None, None, None, None, None

def generate_premium_excel(summary_data, hist_df, div_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: Executive Summary
        df_sum = pd.DataFrame(summary_data)
        df_sum.to_excel(writer, sheet_name="Executive Summary", index=False)

        # Sheet 2: Historical OHLC
        hist_export = hist_df.reset_index()
        if "Date" in hist_export.columns:
            hist_export["Date"] = hist_export["Date"].dt.strftime("%Y-%m-%d")
        hist_export.to_excel(writer, sheet_name="Historical OHLC Data", index=False)

        # Sheet 3: Dividend History
        if not div_df.empty:
            div_export = div_df.reset_index()
            if "Date" in div_export.columns:
                div_export["Date"] = div_export["Date"].dt.strftime("%Y-%m-%d")
            div_export.to_excel(writer, sheet_name="Dividend History", index=False)

        # Styling
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
                ws.column_dimensions[col_letter].width = 28
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

# --- EXECUTION ---
if symbol:
    with st.spinner(get_txt("डेटा लोड हो रहा है...", "Fetching data...")):
        ticker_obj, df_hist, stock_info, ath_val, df_div = fetch_data(symbol, period)

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

        intrinsic_val = calculate_intrinsic_value(eps, book_val) if eps and book_val else None

        # --- DIVIDEND CALCULATIONS ---
        total_lifetime_div = df_div.sum() if not df_div.empty else 0.0
        
        # Yearly Breakdown
        if not df_div.empty:
            df_div_yearly = df_div.groupby(df_div.index.year).sum().reset_index()
            df_div_yearly.columns = ["Year", "Total Dividend"]
            df_div_yearly = df_div_yearly.sort_values(by="Year", ascending=False)
            current_year_div = df_div_yearly[df_div_yearly["Year"] == datetime.date.today().year]["Total Dividend"].sum()
        else:
            df_div_yearly = pd.DataFrame(columns=["Year", "Total Dividend"])
            current_year_div = 0.0

        # Yield on Cost
        yield_on_cost = (div_rate / buy_price * 100) if buy_price > 0 else None

        st.subheader(f"🏢 {long_name} ({symbol})")

        # --- SECTION 1: PRICE & RANGE ---
        st.markdown(f"<div class='sec-header'>{get_txt('मूल्य एवं 52-सप्ताह/लाइफटाइम स्थिति', 'Price Action & ATH Range')}</div>", unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("CMP", f"{currency} {cmp_price:,.2f}", f"{price_change:+.2f} ({price_change_pct:+.2f}%)")
        m2.metric("52W High", f"{currency} {high_52:,.2f}" if high_52 else "N/A", f"{down_from_52w:.2f}% (High)", delta_color="inverse")
        m3.metric("52W Low", f"{currency} {low_52:,.2f}" if low_52 else "N/A")
        m4.metric("Lifetime High (ATH)", f"{currency} {ath:,.2f}" if ath else "N/A", f"{down_from_ath:.2f}% (ATH)", delta_color="inverse")
        if intrinsic_val:
            m5.metric("Intrinsic Value", f"{currency} {intrinsic_val:,.2f}", f"{((intrinsic_val - cmp_price) / cmp_price) * 100:+.1f}% Margin")
        else:
            m5.metric("Intrinsic Value", "N/A")

        # --- SECTION 2: VALUATION ---
        st.markdown(f"<div class='sec-header'>{get_txt('वैल्युएशन एवं फंडामेंटल्स (P/E Multiples)', 'Valuation & Fundamentals')}</div>", unsafe_allow_html=True)
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Stock P/E", f"{company_pe:.2f}" if company_pe else "N/A")
        v2.metric("Industry P/E", str(industry_pe))
        v3.metric("P/B Ratio", f"{pb_ratio:.2f}" if pb_ratio else "N/A")
        v4.metric("EPS (TTM)", f"{currency} {eps:.2f}" if eps else "N/A")
        v5.metric("Dividend Yield (CMP)", f"{div_yield:.2f}%")

        # --- SECTION 3: DIVIDEND INTELLIGENCE & YIELD ON COST ---
        st.markdown(f"<div class='sec-header'>{get_txt('💰 डिविडेंड विश्लेषण एवं पूंजी यील्ड (Dividend Analytics & Yield on Cost)', 'Dividend Analytics & Yield on Cost')}</div>", unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.metric(get_txt("लाइफटाइम कुल डिविडेंड", "Lifetime Total Div"), f"{currency} {total_lifetime_div:,.2f}")
        d2.metric(get_txt("वार्षिक डिविडेंड दर (TTM)", "Annual Div Rate (TTM)"), f"{currency} {div_rate:,.2f}")
        d3.metric(get_txt("वर्तमान वर्ष डिविडेंड", "Current Year Div"), f"{currency} {current_year_div:,.2f}")
        if yield_on_cost is not None:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड (Yield on Cost)", "Yield on Cost (Your Buy)"), f"{yield_on_cost:.2f}%", f"Buy Price: {buy_price}")
        else:
            d4.metric(get_txt("खरीद मूल्य पर यील्ड", "Yield on Cost"), "Sidebar में दर्ज करें")

        # Custom Range Dividend Filter
        with st.expander(get_txt("📅 कस्टम तारीख अनुसार डिविडेंड कैलकुलेटर", "Custom Date Range Dividend Calculator")):
            dcol1, dcol2 = st.columns(2)
            start_d = dcol1.date_input("Start Date", value=datetime.date(2020, 1, 1))
            end_d = dcol2.date_input("End Date", value=datetime.date.today())
            
            if not df_div.empty:
                div_clean = df_div.copy()
                div_clean.index = div_clean.index.tz_localize(None)
                mask = (div_clean.index.date >= start_d) & (div_clean.index.date <= end_d)
                range_div_sum = div_clean[mask].sum()
                st.write(f"**{start_d} से {end_d} के बीच कुल डिविडेंड:** `{currency} {range_div_sum:,.2f}`")
                st.dataframe(div_clean[mask], use_container_width=True)

        # Yearly Table
        if not df_div_yearly.empty:
            with st.expander(get_txt("📊 वर्ष-वार डिविडेंड इतिहास (Yearly Dividend History)", "Yearly Dividend History")):
                st.dataframe(df_div_yearly, use_container_width=True)

        # --- SECTION 4: CHARTS & EXCEL EXPORT ---
        st.markdown(f"<div class='sec-header'>{get_txt('ऐतिहासिक चार्ट', 'Historical Chart')}</div>", unsafe_allow_html=True)
        st.line_chart(df_hist["Close"], use_container_width=True)

        # Prepare Data for Excel
        summary_rows = [
            {"Field": "--- PRICE & ATH METRICS ---", "Value": ""},
            {"Field": "Company Name", "Value": str(long_name)},
            {"Field": "Symbol", "Value": str(symbol)},
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
            {"Field": "--- DIVIDEND & YIELD ON COST ---", "Value": ""},
            {"Field": "Dividend Yield (CMP)", "Value": f"{div_yield:.2f}%"},
            {"Field": "Lifetime Total Dividend", "Value": f"{currency} {total_lifetime_div:,.2f}"},
            {"Field": "Your Buy Price", "Value": f"{currency} {buy_price:,.2f}" if buy_price > 0 else "Not Provided"},
            {"Field": "Yield on Cost (Your Capital)", "Value": f"{yield_on_cost:.2f}%" if yield_on_cost else "N/A"},
        ]

        excel_data = generate_premium_excel(summary_rows, df_hist, df_div)

        st.markdown("---")
        st.download_button(
            label=get_txt("📥 प्रीमियम फॉर्मेटेड एक्सेल रिपोर्ट डाउनलोड करें (.xlsx)", "📥 Download Premium Executive Report (.xlsx)"),
            data=excel_data,
            file_name=f"{symbol}_Executive_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error("डेटा प्राप्त करने में असमर्थ। कृपया सिंबल की जाँच करें।")
