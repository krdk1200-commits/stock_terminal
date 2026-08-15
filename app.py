import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import timedelta, date
import io
import re
import urllib.parse

st.set_page_config(
    page_title="Global Stock & Fundamental Terminal",
    page_icon="🌍",
    layout="wide"
)

# --- सेटिंग्स ---
ADMIN_SECRET_KEY = "DEEPAK108"
MY_UPI_ID = "9661796833@superyes"
PAYEE_NAME = "Deepak Kumar"

TELEGRAM_BOT_TOKEN = ""
TELEGRAM_CHAT_ID = ""

def send_telegram_alert(text_message):
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        try:
            import requests
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
            payload = {"chat_id": TELEGRAM_CHAT_ID, "text": text_message, "parse_mode": "Markdown"}
            requests.post(url, json=payload, timeout=5)
        except Exception:
            pass

def sanitize_input(text: str) -> str:
    if not text: return ""
    return re.sub(r'[<>{}\;]', '', text).strip()[:500]

# प्रमुख भारतीय और अमेरिकी इंडेक्स
GLOBAL_INDICES = {
    "🇮🇳 NIFTY 50": "^NSEI",
    "🇮🇳 BANK NIFTY": "^NSEBANK",
    "🇮🇳 NIFTY IT": "^CNXIT",
    "🇮🇳 BSE SENSEX": "^BSESN",
    "🇺🇸 S&P 500": "^GSPC",
    "🇺🇸 NASDAQ 100": "^NDX",
    "🇺🇸 DOW JONES": "^DJI",
    "🇺🇸 RUSSELL 2000": "^RUT"
}

@st.cache_data(ttl=60)
def fetch_stock_data(ticker_sym):
    t = yf.Ticker(ticker_sym)
    hist = t.history(period="max")
    inf = t.info if hasattr(t, 'info') and t.info else {}
    inc = t.financials if hasattr(t, 'financials') else pd.DataFrame()
    bal = t.balance_sheet if hasattr(t, 'balance_sheet') else pd.DataFrame()
    cf = t.cashflow if hasattr(t, 'cashflow') else pd.DataFrame()
    return hist, inf, inc, bal, cf

# --- साइडबार ---
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/bullish.png", width=70)
    st.markdown("### ⚙️ सेटिंग्स / Settings")
    lang_choice = st.radio(
        "🌐 भाषा चुनें / Select Language:",
        ["Bilingual (हिंदी + English)", "हिंदी (Hindi)", "English"]
    )
    st.markdown("---")
    st.markdown("### 🌍 लोकप्रिय इंडेक्स / Global Indices")
    selected_sidebar_index = st.selectbox(
        "इंडेक्स चुनें / Select Index:",
        ["-- Manual / सिंबल दर्ज करें --"] + list(GLOBAL_INDICES.keys())
    )
    st.markdown("---")
    st.markdown("### 🔐 एडमिन अनलॉक (Admin Access)")
    admin_input_key = st.text_input("एडमिन पासकोड दर्ज करें:", type="password")
    is_admin = (admin_input_key == ADMIN_SECRET_KEY)
    if is_admin:
        st.success("✅ एडमिन एक्सेस एक्टिव (Free Direct Download)")

# --- टॉप एडवर्टाइजमेंट ---
st.markdown("""
    <div style="background: linear-gradient(90deg, #1e293b, #0f172a); border: 1px dashed #3b82f6; padding: 10px; border-radius: 8px; text-align: center; margin-bottom: 15px;">
        <span style="color: #94a3b8; font-size: 12px; text-transform: uppercase; letter-spacing: 1px;">📢 Sponsored / Advertisement</span>
        <p style="margin: 5px 0 0 0; color: #38bdf8; font-weight: 600; font-size: 14px;">
            ⚡ Zero Brokerage Global & Indian Stock Investing | Open Account Now
        </p>
    </div>
""", unsafe_allow_html=True)

texts = {
    "English": {
        "title": "🌍 Global Stock & Index Market Terminal",
        "sub": "Analyze Indian & US Stocks/Indices. Full Date-wise Records, Pivots, Charts & Company Fundamentals.",
        "sym_label": "Enter ANY Indian or US Stock/Index (e.g. RELIANCE, TCS, AAPL, TSLA, NVDA, NIFTY, ^GSPC):",
        "mode_label": "Date Range Mode:",
        "btn": "🚀 Analyze & Display Data",
        "loading": "Connecting to Global Exchanges...",
        "tab_daily": "📅 Day-by-Day Datewise Log",
        "tab_chart": "📈 Price Chart View",
        "tab_summary": "📊 Timeframe Summary",
        "tab_pivot": "📑 Pivot Tables",
        "tab_fund": "🏢 Fundamentals & Cash",
        "tab_financials": "📈 Financial Statements",
        "sugg_head": "💡 Help Me to Improve This Website",
        "sugg_sub": "Submit feedback or feature requests directly to the creator!",
        "sugg_btn": "Submit Suggestion",
        "sugg_success": "Thank you! Your suggestion has been recorded."
    },
    "हिंदी (Hindi)": {
        "title": "🌍 ग्लोबल स्टॉक एवं इंडेक्स मार्केट टर्मिनल",
        "sub": "भारतीय और अमेरिकी स्टॉक्स/इंडेक्स का संपूर्ण विश्लेषण: दैनिक रिकॉर्ड, चार्ट्स, पिवट और फंडामेंटल्स।",
        "sym_label": "किसी भी भारतीय या अमेरिकी स्टॉक/इंडेक्स का नाम दर्ज करें (उदा. RELIANCE, TCS, AAPL, TSLA, NVDA, NIFTY, ^GSPC):",
        "mode_label": "समयावधि चयन मोड:",
        "btn": "🚀 संपूर्ण डेटा निकालें एवं विश्लेषण देखें",
        "loading": "ग्लोबल मार्केट एक्सचेंज से डेटा लोड हो रहा है...",
        "tab_daily": "📅 दिन-प्रतिदिन (तारीख अनुसार) दैनिक रिकॉर्ड",
        "tab_chart": "📈 प्राइस चार्ट व्यू",
        "tab_summary": "📊 टाइमफ्रेम सारांश",
        "tab_pivot": "📑 पिवट टेबल्स",
        "tab_fund": "🏢 फंडामेंटल्स एवं कैश",
        "tab_financials": "📈 वित्तीय विवरण",
        "sugg_head": "💡 वेबसाइट को बेहतर बनाने में मदद करें (Help Me to Improve This Website)",
        "sugg_sub": "इस वेबसाइट को और बेहतर बनाने के लिए अपना सुझाव भेजें!",
        "sugg_btn": "सुझाव सबमिट करें",
        "sugg_success": "धन्यवाद! आपका सुझाव सुरक्षित रूप से दर्ज कर लिया गया है।"
    },
    "Bilingual (हिंदी + English)": {
        "title": "🌍 Global Stock Terminal | वैश्विक मार्केट टर्मिनल",
        "sub": "Indian & US Stocks / Indices Analysis | भारतीय एवं अमेरिकी स्टॉक्स विश्लेषण",
        "sym_label": "Stock / Index Symbol (भारतीय या अमेरिकी सिंबल):",
        "mode_label": "Range Type / मोड:",
        "btn": "🚀 Analyze / डेटा निकालें",
        "loading": "Fetching Market Data...",
        "tab_daily": "📅 Datewise Daily Log / दैनिक रिकॉर्ड",
        "tab_chart": "📈 Chart / चार्ट",
        "tab_summary": "📊 Summary / सारांश",
        "tab_pivot": "📑 Pivot Tables / पिवट टेबल",
        "tab_fund": "🏢 Fundamentals / फंडामेंटल्स",
        "tab_financials": "📈 Statements / वित्तीय विवरण",
        "sugg_head": "💡 Help Me to Improve This Website | सुझाव दें",
        "sugg_sub": "Share suggestions to improve the terminal / अपनी राय यहाँ भेजें",
        "sugg_btn": "Submit / भेजें",
        "sugg_success": "Suggestion received! / आपका सुझाव प्राप्त हुआ!"
    }
}

T = texts[lang_choice]

st.title(T["title"])
st.caption(T["sub"])
st.markdown("---")

default_input = "NIFTY"
if selected_sidebar_index != "-- Manual / सिंबल दर्ज करें --":
    default_input = selected_sidebar_index

c_sym, c_mode = st.columns([3, 2])
with c_sym:
    user_symbol = st.text_input(T["sym_label"], value=default_input)

with c_mode:
    range_mode = st.radio(
        T["mode_label"],
        ["Standard Presets (1D, 1W, 1M, 1Y...)", "📅 Custom Date Range (कैलेंडर से तारीख चुनें)"],
        horizontal=True
    )

start_date_filter, end_date_filter = None, None
selected_dur_days = 30

if range_mode == "Standard Presets (1D, 1W, 1M, 1Y...)":
    duration_options = {
        "1 Day / 1 दिन": 1, "2 Days / 2 दिन": 2, "1 Week (7 Days) / 1 सप्ताह": 7,
        "4 Weeks (28 Days) / 4 सप्ताह": 28, "5 Weeks (35 Days) / 5 सप्ताह": 35,
        "1 Month (30 Days) / 1 माह": 30, "60 Days / 60 दिन": 60,
        "3 Months (90 Days) / 3 माह": 90, "6 Months (180 Days) / 6 माह": 180,
        "1 Year (365 Days) / 1 वर्ष": 365, "5 Years (1825 Days) / 5 वर्ष": 1825,
        "10 Years (3650 Days) / 10 वर्ष": 3650, "All-Time History / शुरुआत से अब तक": 999999
    }
    selected_preset = st.selectbox("समयावधि चुनें / Duration:", list(duration_options.keys()), index=5)
    selected_dur_days = duration_options[selected_preset]
else:
    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date_filter = st.date_input("Start Date:", value=date.today() - timedelta(days=90))
    with col_d2: end_date_filter = st.date_input("End Date:", value=date.today())

# कस्टमाइज़ेशन सेटिंग्स
with st.expander("🛠️ कस्टमाइज़ेशन विकल्प (Custom Columns & Sheets)", expanded=False):
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        available_cols = ["Open", "High", "Low", "Close", "Typical Avg Price", "20-Day SMA", "Points Diff", "Daily Change (%)", "Volume"]
        selected_cols = st.multiselect("चयनित कॉलम:", available_cols, default=["Open", "High", "Low", "Close", "Typical Avg Price", "20-Day SMA", "Points Diff", "Daily Change (%)"])
    with col_k2:
        available_sheets = ["Datewise_Daily_Log", "Timeframe_Summary", "Yearly_Pivot", "Monthly_Pivot", "Key_Fundamentals", "Income_Statement", "Balance_Sheet", "Cash_Flow"]
        selected_sheets = st.multiselect("डाउनलोड हेतु शीट्स:", available_sheets, default=["Datewise_Daily_Log", "Timeframe_Summary", "Yearly_Pivot", "Monthly_Pivot", "Key_Fundamentals"])

# ट्रिगर
if st.button(T["btn"], use_container_width=True):
    with st.spinner(T["loading"]):
        raw_sym = user_symbol.strip().upper()

        # अमेरिकी और भारतीय सिंबल रेजोल्यूशन
        if raw_sym in GLOBAL_INDICES:
            ticker_symbol = GLOBAL_INDICES[raw_sym]
        elif raw_sym in ["NIFTY", "NIFTY50", "NIFTY 50"]: ticker_symbol = "^NSEI"
        elif raw_sym in ["BANKNIFTY", "BANK NIFTY"]: ticker_symbol = "^NSEBANK"
        elif raw_sym in ["SENSEX", "BSE"]: ticker_symbol = "^BSESN"
        elif raw_sym in ["SP500", "S&P500"]: ticker_symbol = "^GSPC"
        elif raw_sym in ["NASDAQ", "NDX"]: ticker_symbol = "^NDX"
        elif raw_sym in ["DOW", "DJI"]: ticker_symbol = "^DJI"
        else:
            # चेक करें कि यह अमेरिकी स्टॉक है (जैसे AAPL, TSLA, NVDA) या भारतीय स्टॉक
            # पहले डायरेक्ट सिंबल चेक करें
            t_test = yf.Ticker(raw_sym)
            df_test = t_test.history(period="5d")
            if not df_test.empty:
                ticker_symbol = raw_sym
            elif not raw_sym.startswith('^') and not raw_sym.endswith('.NS') and not raw_sym.endswith('.BO'):
                ticker_symbol = f"{raw_sym}.NS"
            else:
                ticker_symbol = raw_sym

        try:
            df, info, inc_stmt, bal_sheet, cash_flow = fetch_stock_data(ticker_symbol)
            if df.empty and ticker_symbol.endswith('.NS'):
                ticker_symbol = ticker_symbol.replace('.NS', '.BO')
                df, info, inc_stmt, bal_sheet, cash_flow = fetch_stock_data(ticker_symbol)

            if df.empty:
                st.error(f"❌ '{raw_sym}' का डेटा नहीं मिला। कृपया सिंबल की जांच करें (उदा. AAPL, TSLA, NVDA, RELIANCE, TCS)।")
            else:
                company_name = info.get("longName") or info.get("shortName") or raw_sym
                curr_symbol = "$" if info.get("currency") == "USD" else "₹"
                sector = info.get("sector", "Index / Market Segment")
                st.success(f"🏢 **{company_name}** | सिंबल: `{ticker_symbol}` | मुद्रा: `{info.get('currency', 'INR')}` | सेक्टर: *{sector}*")

                last_date = df.index[-1]

                # 1. डेटा कैलकुलेशन
                df_work = df.copy()
                df_work['Typical_Avg'] = round((df_work['High'] + df_work['Low'] + df_work['Close']) / 3, 2)
                df_work['SMA_20'] = round(df_work['Close'].rolling(window=20, min_periods=1).mean(), 2)
                df_work['Prev_Close'] = df_work['Close'].shift(1)
                df_work['Daily_Diff'] = round(df_work['Close'] - df_work['Prev_Close'], 2).fillna(0)
                df_work['Daily_Pct'] = round(((df_work['Close'] - df_work['Prev_Close']) / df_work['Prev_Close']) * 100, 2).fillna(0)

                if range_mode == "Standard Presets (1D, 1W, 1M, 1Y...)":
                    if selected_dur_days == 1: day_df = df_work.iloc[-1:].copy()
                    elif selected_dur_days == 2: day_df = df_work.iloc[-2:].copy()
                    elif selected_dur_days != 999999: day_df = df_work[df_work.index >= (last_date - timedelta(days=selected_dur_days))].copy()
                    else: day_df = df_work.copy()
                else:
                    day_df = df_work[(df_work.index.date >= start_date_filter) & (df_work.index.date <= end_date_filter)].copy()

                chart_df = day_df.copy() # चार्ट के लिए
                day_df = day_df.sort_index(ascending=False)

                base_dict = {
                    "Date (तारीख)": day_df.index.strftime('%d-%m-%Y'),
                    "Open": round(day_df['Open'], 2),
                    "High": round(day_df['High'], 2),
                    "Low": round(day_df['Low'], 2),
                    "Close": round(day_df['Close'], 2),
                    "Typical Avg Price": day_df['Typical_Avg'],
                    "20-Day SMA": day_df['SMA_20'],
                    "Points Diff": day_df['Daily_Diff'],
                    "Daily Change (%)": day_df['Daily_Pct'].apply(lambda x: f"{x:+.2f}%")
                }
                if 'Volume' in day_df.columns:
                    base_dict["Volume"] = day_df['Volume']

                full_detailed_df = pd.DataFrame(base_dict)
                chosen_cols = ["Date (तारीख)"] + [c for c in selected_cols if c in full_detailed_df.columns]
                detailed_df = full_detailed_df[chosen_cols]

                # 2. टाइमफ्रेम समरी
                timeframe_intervals = [
                    ("1 Day", 1), ("1 Week", 7), ("4 Weeks", 28), ("5 Weeks", 35),
                    ("1 Month", 30), ("3 Months", 90), ("6 Months", 180),
                    ("1 Year", 365), ("5 Years", 1825), ("10 Years", 3650), ("All-Time (Max)", len(df))
                ]
                sum_list = []
                for label, days in timeframe_intervals:
                    sub = df.iloc[-1:] if days == 1 else (df if days == len(df) else df[df.index >= (last_date - timedelta(days=days))])
                    if not sub.empty:
                        o, h, l, c = round(sub['Open'].iloc[0], 2), round(sub['High'].max(), 2), round(sub['Low'].min(), 2), round(sub['Close'].iloc[-1], 2)
                        sma = round(sub['Close'].mean(), 2)
                        avg_p = round(((sub['High'] + sub['Low'] + sub['Close']) / 3).mean(), 2)
                        diff = round(c - o, 2)
                        pct = round(((c - o) / o) * 100, 2)
                        sum_list.append({"Timeframe": label, "Open": o, "High": h, "Low": l, "Close": c, "Avg Price": avg_p, "SMA / Moving Avg": sma, "Point Diff": diff, "Change (%)": f"{pct:+.2f}%"})
                summary_df = pd.DataFrame(sum_list)

                # 3. पिवट टेबल
                df_pivot = df.copy()
                df_pivot['Year'] = df_pivot.index.year
                df_pivot['Month_Name'] = df_pivot.index.strftime('%B')
                df_pivot['Typical_Price'] = (df_pivot['High'] + df_pivot['Low'] + df_pivot['Close']) / 3

                yearly_pivot = df_pivot.groupby('Year').agg(
                    Year_Open=('Open', 'first'), Year_High=('High', 'max'), Year_Low=('Low', 'min'), Year_Close=('Close', 'last'), Year_Avg_Price=('Typical_Price', 'mean')
                ).round(2)
                yearly_pivot['Yearly_Return (%)'] = (((yearly_pivot['Year_Close'] - yearly_pivot['Year_Open']) / yearly_pivot['Year_Open']) * 100).round(2).apply(lambda x: f"{x:+.2f}%")
                yearly_pivot = yearly_pivot.sort_index(ascending=False).reset_index()

                monthly_pivot = pd.pivot_table(df_pivot, values='Typical_Price', index='Year', columns='Month_Name', aggfunc='mean').round(2).sort_index(ascending=False).fillna("-")

                # 4. फंडामेंटल्स
                def fmt_num(val):
                    if val is None or pd.isna(val): return "N/A"
                    if isinstance(val, (int, float)):
                        if curr_symbol == "₹" and abs(val) >= 1e7:
                            return f"₹ {val / 1e7:,.2f} Cr"
                        elif curr_symbol == "$" and abs(val) >= 1e9:
                            return f"$ {val / 1e9:,.2f} B"
                        return f"{curr_symbol} {val:,.2f}"
                    return str(val)

                fund_df = pd.DataFrame([
                    {"Metric": "Company Name", "Value": company_name},
                    {"Metric": "Market Cap", "Value": fmt_num(info.get("marketCap"))},
                    {"Metric": "Current Price", "Value": f"{curr_symbol} {info.get('currentPrice', df['Close'].iloc[-1]):,.2f}"},
                    {"Metric": "P/E Ratio", "Value": info.get("trailingPE", "N/A")},
                    {"Metric": "EV / EBITDA", "Value": info.get("enterpriseToEbitda", "N/A")},
                    {"Metric": "EBITDA", "Value": fmt_num(info.get("ebitda"))},
                    {"Metric": "Cash Reserves", "Value": fmt_num(info.get("totalCash"))},
                    {"Metric": "Total Debt", "Value": fmt_num(info.get("totalDebt"))},
                    {"Metric": "Debt to Equity", "Value": f"{info.get('debtToEquity', 'N/A')}"},
                    {"Metric": "ROE", "Value": f"{info.get('returnOnEquity', 0)*100:.2f}%" if info.get('returnOnEquity') else "N/A"}
                ])

                # एक्सेल एक्सपोर्ट बफर
                custom_excel = io.BytesIO()
                with pd.ExcelWriter(custom_excel, engine='openpyxl') as writer:
                    if "Datewise_Daily_Log" in selected_sheets: detailed_df.to_excel(writer, index=False, sheet_name='Datewise_Daily_Log')
                    if "Timeframe_Summary" in selected_sheets: summary_df.to_excel(writer, index=False, sheet_name='Timeframe_Summary')
                    if "Yearly_Pivot" in selected_sheets: yearly_pivot.to_excel(writer, index=False, sheet_name='Yearly_Pivot')
                    if "Monthly_Pivot" in selected_sheets: monthly_pivot.to_excel(writer, sheet_name='Monthly_Pivot')
                    if "Key_Fundamentals" in selected_sheets: fund_df.to_excel(writer, index=False, sheet_name='Key_Fundamentals')
                    if "Income_Statement" in selected_sheets and not inc_stmt.empty: inc_stmt.to_excel(writer, sheet_name='Income_Statement')
                    if "Balance_Sheet" in selected_sheets and not bal_sheet.empty: bal_sheet.to_excel(writer, sheet_name='Balance_Sheet')
                    if "Cash_Flow" in selected_sheets and not cash_flow.empty: cash_flow.to_excel(writer, sheet_name='Cash_Flow')

                # --- 📥 डाउनलोड सेक्शन ---
                st.markdown("### 📥 कस्टमाइज्ड एक्सेल डाउनलोड (Download Excel)")

                if is_admin:
                    st.download_button(
                        label=f"👑 [Admin Direct Free Download] {raw_sym}_Report.xlsx",
                        data=custom_excel.getvalue(),
                        file_name=f"{raw_sym}_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    with st.expander("⚡ ₹10 शुल्क देकर फाइल डाउनलोड करें (Click to Unlock)", expanded=True):
                        col_pay1, col_pay2 = st.columns([1, 2])
                        with col_pay1:
                            tip_amount = st.selectbox("राशि चुनें / Amount:", [10, 20, 50, 100], index=0, format_func=lambda x: f"₹{x} (Standard)" if x==10 else f"₹{x} (💖 Super Thanks Tip)")
                            upi_url = f"upi://pay?pa={MY_UPI_ID}&pn={urllib.parse.quote(PAYEE_NAME)}&am={tip_amount}&cu=INR&tn={raw_sym}_Global_Report"
                            qr_api_url = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data={urllib.parse.quote(upi_url)}"
                            st.image(qr_api_url, caption=f"Scan & Pay ₹{tip_amount}")

                        with col_pay2:
                            st.markdown(f"""
                            **डाउनलोड करने के स्टेप्स:**
                            1. QR कोड स्कैन करके **₹{tip_amount}** का भुगतान करें।
                            2. भुगतान पूरा होने के बाद नीचे दिए गए चेकबॉक्स पर टिक करें।
                            """)
                            if st.checkbox("मैंने भुगतान कर दिया है / I Have Paid"):
                                st.download_button(
                                    label=f"📥 Download {raw_sym}_Master_Report.xlsx",
                                    data=custom_excel.getvalue(),
                                    file_name=f"{raw_sym}_Master_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )

                st.markdown("---")

                # टैब्स
                tab1, tab_chart, tab2, tab3, tab4, tab5 = st.tabs([
                    T["tab_daily"], T["tab_chart"], T["tab_summary"], T["tab_pivot"], T["tab_fund"], T["tab_financials"]
                ])
                with tab1:
                    st.markdown(f"### 📅 दैनिक रिकॉर्ड ({len(detailed_df)} ट्रेडिंग दिन)")
                    st.dataframe(detailed_df, use_container_width=True, height=450)
                with tab_chart:
                    st.markdown(f"### 📈 {raw_sym} क्लोजिंग प्राइस एवं 20-Day SMA चार्ट")
                    st.line_chart(chart_df[['Close', 'SMA_20']])
                with tab2:
                    st.dataframe(summary_df, use_container_width=True)
                with tab3:
                    st.write("**वार्षिक पिवट:**")
                    st.dataframe(yearly_pivot, use_container_width=True)
                    st.write("**मासिक पिवट:**")
                    st.dataframe(monthly_pivot, use_container_width=True)
                with tab4:
                    st.dataframe(fund_df, use_container_width=True)
                with tab5:
                    if not inc_stmt.empty: st.dataframe(inc_stmt, use_container_width=True)

        except Exception as ex:
            st.error(f"Error: {ex}")

# --- 💡 HELP ME TO IMPROVE THIS WEBSITE ---
st.markdown("---")
col_logo, col_title = st.columns([1, 11])
with col_logo:
    st.image("https://img.icons8.com/fluency/96/idea.png", width=60)
with col_title:
    st.markdown(f"### {T['sugg_head']}")
    st.caption(T["sugg_sub"])

with st.form(key="telegram_suggestion_form"):
    sugg_user = st.text_input("आपका नाम / संपर्क (वैकल्पिक):")
    sugg_text = st.text_area("इस वेबसाइट में सुधार या नया फीचर सुझाव (Your Suggestion):", height=90)
    send_btn = st.form_submit_button(T["sugg_btn"])

    if send_btn:
        clean_sugg = sanitize_input(sugg_text)
        clean_user = sanitize_input(sugg_user) or "Anonymous User"

        if clean_sugg:
            with open("suggestions.log", "a", encoding="utf-8") as f:
                f.write(f"User: {clean_user} | Suggestion: {clean_sugg}\n")

            msg_body = f"🚀 *नया सुझाव प्राप्त हुआ!*\n\n👤 *यूज़र:* {clean_user}\n📝 *सुझाव:* {clean_sugg}"
            send_telegram_alert(msg_body)
            st.success(T["sugg_success"])
        else:
            st.warning("कृपया अपना सुझाव लिखें।")

# बॉटम बैनर
st.markdown("""
    <div style="background: #0f172a; border-top: 1px solid #334155; padding: 12px; border-radius: 8px; text-align: center; margin-top: 25px;">
        <span style="color: #64748b; font-size: 11px; text-transform: uppercase;">Global Financial Terminal</span>
        <p style="margin: 3px 0 0 0; color: #94a3b8; font-size: 13px;">
            📊 Indian & US Equity Analytics Engine | All Rights Reserved © 2026
        </p>
    </div>
""", unsafe_allow_html=True)
