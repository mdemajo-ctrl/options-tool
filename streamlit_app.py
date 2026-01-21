"""
Options Pricing & P&L Tool - Streamlit Web App
Deploy to Streamlit Cloud for free sharing.
"""

import streamlit as st
import pandas as pd
import yfinance as yf
import math
from datetime import datetime
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# Try scipy, fall back to math.erf
try:
    from scipy.stats import norm
    def normal_cdf(x):
        return norm.cdf(x)
except ImportError:
    def normal_cdf(x):
        return 0.5 * (1 + math.erf(x / math.sqrt(2)))

# Constants
RISK_FREE_RATE = 0.045

st.set_page_config(page_title="Options Pricing Tool", page_icon="📈", layout="wide")

st.title("📈 Options Pricing & P&L Tool")


def calculate_delta(spot, strike, time_to_expiry, volatility, option_type="CALL"):
    if time_to_expiry <= 0 or volatility <= 0 or spot <= 0 or strike <= 0:
        return 0.5 if option_type == "CALL" else -0.5
    try:
        d1 = (math.log(spot / strike) + (RISK_FREE_RATE + 0.5 * volatility ** 2) * time_to_expiry) / (volatility * math.sqrt(time_to_expiry))
        return normal_cdf(d1) if option_type == "CALL" else normal_cdf(d1) - 1
    except:
        return 0.5 if option_type == "CALL" else -0.5


@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_stock_data(ticker):
    """Fetch stock price and available expirations."""
    stock = yf.Ticker(ticker)
    try:
        expirations = list(stock.options)
    except:
        return None, [], 0

    try:
        price = stock.info.get('regularMarketPrice') or stock.info.get('currentPrice') or 0
        if not price:
            hist = stock.history(period="1d")
            price = hist['Close'].iloc[-1] if not hist.empty else 0
    except:
        price = 0

    return stock, expirations, price


@st.cache_data(ttl=300)
def get_options_chain(ticker, expiry):
    """Fetch options chain for given expiry."""
    stock = yf.Ticker(ticker)
    try:
        chain = stock.option_chain(expiry)
        return chain.calls, chain.puts
    except:
        return None, None


def process_options_df(df, current_price, time_to_expiry, option_type):
    """Process options dataframe and calculate delta."""
    result = []
    for _, row in df.iterrows():
        strike = row.get('strike', 0)
        bid = row.get('bid', 0) or 0
        ask = row.get('ask', 0) or 0
        last = row.get('lastPrice', 0) or 0
        mid = (bid + ask) / 2 if bid > 0 and ask > 0 else last
        volume = row.get('volume', 0) or 0
        oi = row.get('openInterest', 0) or 0
        iv = row.get('impliedVolatility', 0) or 0

        delta = calculate_delta(current_price, strike, time_to_expiry, iv, option_type) if iv > 0 else 0

        result.append({
            'Strike': strike,
            'Bid': bid,
            'Ask': ask,
            'Last': last,
            'Mid': mid,
            'Volume': int(volume) if pd.notna(volume) else 0,
            'Open Int': int(oi) if pd.notna(oi) else 0,
            'Impl Vol': iv,
            'Delta': delta,
        })

    return pd.DataFrame(result)


def create_excel_download(ticker, calls_df, puts_df, expiry, current_price, expirations):
    """Create Excel file for download."""
    DARK_HEADER_FILL = PatternFill(start_color="44546A", end_color="44546A", fill_type="solid")
    BLUE_HEADER_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    YELLOW_INPUT_FILL = PatternFill(start_color="FFFFC8", end_color="FFFFC8", fill_type="solid")
    WHITE_FONT = Font(color="FFFFFF", bold=True)
    THIN_BORDER = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def create_sheet(ws, options_df, option_type):
        expiry_date = datetime.strptime(expiry, "%Y-%m-%d")
        days_to_expiry = (expiry_date - datetime.now()).days
        time_to_expiry = max(days_to_expiry / 365.0, 0.001)

        # Row 1 headers
        for col, text in [(2, "Ticker"), (3, "Expiry"), (4, "Days"), (6, "Current Price"), (8, "Stock @ Expiry")]:
            cell = ws.cell(row=1, column=col, value=text)
            cell.fill = DARK_HEADER_FILL if col <= 6 else BLUE_HEADER_FILL
            cell.font = WHITE_FONT
            cell.border = THIN_BORDER

        # Row 2 values
        ws.cell(row=2, column=2, value=ticker).border = THIN_BORDER
        ws.cell(row=2, column=3, value=expiry).border = THIN_BORDER
        ws.cell(row=2, column=4, value=days_to_expiry).border = THIN_BORDER
        ws.cell(row=2, column=6, value=current_price).number_format = '$#,##0.00'
        ws.cell(row=2, column=6).border = THIN_BORDER

        exp_cell = ws.cell(row=2, column=8, value=current_price)
        exp_cell.number_format = '$#,##0.00'
        exp_cell.fill = YELLOW_INPUT_FILL
        exp_cell.border = THIN_BORDER

        # Stock position inputs
        ws.cell(row=1, column=10, value="Stock Shares").fill = BLUE_HEADER_FILL
        ws.cell(row=1, column=10).font = WHITE_FONT
        ws.cell(row=1, column=10).border = THIN_BORDER
        ws.cell(row=1, column=11, value="Stock Entry").fill = BLUE_HEADER_FILL
        ws.cell(row=1, column=11).font = WHITE_FONT
        ws.cell(row=1, column=11).border = THIN_BORDER
        ws.cell(row=1, column=12, value="Stock P&L").fill = BLUE_HEADER_FILL
        ws.cell(row=1, column=12).font = WHITE_FONT
        ws.cell(row=1, column=12).border = THIN_BORDER

        ws.cell(row=2, column=10, value=0).fill = YELLOW_INPUT_FILL
        ws.cell(row=2, column=10).border = THIN_BORDER
        ws.cell(row=2, column=11, value=current_price).number_format = '$#,##0.00'
        ws.cell(row=2, column=11).fill = YELLOW_INPUT_FILL
        ws.cell(row=2, column=11).border = THIN_BORDER
        ws.cell(row=2, column=12, value="=(H2-K2)*J2").number_format = '$#,##0.00'
        ws.cell(row=2, column=12).border = THIN_BORDER
        ws.cell(row=2, column=12).font = Font(bold=True)

        # Summary row 3
        ws.cell(row=3, column=1, value="SUMMARY").font = Font(bold=True, size=11)
        ws.cell(row=3, column=2, value="Premiums Paid:").font = Font(bold=True)
        ws.cell(row=3, column=4, value="Premiums Rcvd:").font = Font(bold=True)
        ws.cell(row=3, column=6, value="Options Payout:").font = Font(bold=True)
        ws.cell(row=3, column=8, value="Options P&L:").font = Font(bold=True)
        ws.cell(row=3, column=10, value="Stock P&L:").font = Font(bold=True)
        ws.cell(row=3, column=11, value="=L2").number_format = '$#,##0.00'
        ws.cell(row=3, column=11).font = Font(bold=True)
        ws.cell(row=3, column=12, value="TOTAL P&L:").font = Font(bold=True, color="0070C0")

        # Column headers row 5
        headers = ["Strike", "Bid", "Ask", "Last", "Mid", "Volume", "Open Int", "Impl Vol", "Delta",
                   "Position", "Entry", "Prem Paid", "Prem Rcvd", "Val@Exp", "Payout", "P&L"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=h)
            cell.fill = DARK_HEADER_FILL if col <= 9 else BLUE_HEADER_FILL
            cell.font = WHITE_FONT
            cell.border = THIN_BORDER

        # Data rows
        row = 6
        for _, r in options_df.iterrows():
            strike = r['Strike']
            mid = r['Mid']
            iv = r['Impl Vol']
            delta = r['Delta']

            ws.cell(row=row, column=1, value=strike).number_format = '$#,##0.00'
            ws.cell(row=row, column=2, value=r['Bid']).number_format = '$#,##0.00'
            ws.cell(row=row, column=3, value=r['Ask']).number_format = '$#,##0.00'
            ws.cell(row=row, column=4, value=r['Last']).number_format = '$#,##0.00'
            ws.cell(row=row, column=5, value=mid).number_format = '$#,##0.00'
            ws.cell(row=row, column=6, value=r['Volume']).number_format = '#,##0'
            ws.cell(row=row, column=7, value=r['Open Int']).number_format = '#,##0'
            ws.cell(row=row, column=8, value=iv).number_format = '0.00%'
            ws.cell(row=row, column=9, value=delta).number_format = '0.00'

            ws.cell(row=row, column=10, value=0).fill = YELLOW_INPUT_FILL
            ws.cell(row=row, column=11, value=mid).number_format = '$#,##0.00'
            ws.cell(row=row, column=11).fill = YELLOW_INPUT_FILL

            ws.cell(row=row, column=12, value=f"=IF(J{row}>0,K{row}*J{row}*100,0)").number_format = '$#,##0.00'
            ws.cell(row=row, column=13, value=f"=IF(J{row}<0,K{row}*-J{row}*100,0)").number_format = '$#,##0.00'

            if option_type == "CALL":
                ws.cell(row=row, column=14, value=f"=MAX($H$2-A{row},0)").number_format = '$#,##0.00'
            else:
                ws.cell(row=row, column=14, value=f"=MAX(A{row}-$H$2,0)").number_format = '$#,##0.00'

            ws.cell(row=row, column=15, value=f"=N{row}*J{row}*100").number_format = '$#,##0.00'
            ws.cell(row=row, column=16, value=f"=O{row}-L{row}+M{row}").number_format = '$#,##0.00'

            row += 1

        last_row = row - 1

        # Summary formulas
        ws.cell(row=3, column=3, value=f"=SUM(L6:L{last_row})").number_format = '$#,##0.00'
        ws.cell(row=3, column=3).font = Font(bold=True)
        ws.cell(row=3, column=5, value=f"=SUM(M6:M{last_row})").number_format = '$#,##0.00'
        ws.cell(row=3, column=5).font = Font(bold=True)
        ws.cell(row=3, column=7, value=f"=SUM(O6:O{last_row})").number_format = '$#,##0.00'
        ws.cell(row=3, column=7).font = Font(bold=True)
        ws.cell(row=3, column=9, value=f"=SUM(P6:P{last_row})").number_format = '$#,##0.00'
        ws.cell(row=3, column=9).font = Font(bold=True)
        ws.cell(row=3, column=13, value="=I3+K3").number_format = '$#,##0.00'
        ws.cell(row=3, column=13).font = Font(bold=True, size=12, color="0070C0")

        # Conditional formatting
        green_font = Font(color="008000")
        red_font = Font(color="C00000")
        ws.conditional_formatting.add(f"P6:P{last_row}", CellIsRule(operator='greaterThan', formula=['0'], font=green_font))
        ws.conditional_formatting.add(f"P6:P{last_row}", CellIsRule(operator='lessThan', formula=['0'], font=red_font))

        # Column widths
        for i, w in enumerate([10, 8, 8, 8, 8, 9, 9, 9, 7, 8, 8, 10, 10, 9, 10, 10], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A6'

    wb = Workbook()
    ws_calls = wb.active
    ws_calls.title = "Calls"
    create_sheet(ws_calls, calls_df, "CALL")

    ws_puts = wb.create_sheet("Puts")
    create_sheet(ws_puts, puts_df, "PUT")

    # Expirations tab
    ws_exp = wb.create_sheet("Expirations")
    ws_exp.cell(row=1, column=1, value="Available Expirations").font = Font(bold=True)
    for i, exp in enumerate(expirations, 3):
        cell = ws_exp.cell(row=i, column=1, value=exp)
        if exp == expiry:
            cell.font = Font(bold=True, color="0070C0")

    # Save to bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- Main App UI ---

# Sidebar for inputs
with st.sidebar:
    st.header("Settings")

    ticker = st.text_input("Ticker Symbol", value="CCJ").upper()

    if st.button("Load Expirations", type="primary"):
        with st.spinner(f"Loading {ticker}..."):
            stock, expirations, price = get_stock_data(ticker)
            if expirations:
                st.session_state['expirations'] = expirations
                st.session_state['current_price'] = price
                st.session_state['ticker'] = ticker
                st.success(f"Found {len(expirations)} expirations")
            else:
                st.error("No options found for this ticker")

    # Expiration dropdown
    if 'expirations' in st.session_state and st.session_state['expirations']:
        expiry = st.selectbox("Expiration Date", st.session_state['expirations'])
        st.metric("Current Price", f"${st.session_state.get('current_price', 0):.2f}")

        # Stock @ Expiry input
        stock_at_expiry = st.number_input(
            "Stock @ Expiry",
            value=st.session_state.get('current_price', 100.0),
            step=1.0,
            format="%.2f"
        )
        st.session_state['stock_at_expiry'] = stock_at_expiry
        st.session_state['selected_expiry'] = expiry

# Main content
if 'expirations' in st.session_state and st.session_state['expirations']:
    ticker = st.session_state.get('ticker', '')
    expiry = st.session_state.get('selected_expiry', '')
    current_price = st.session_state.get('current_price', 0)
    stock_at_expiry = st.session_state.get('stock_at_expiry', current_price)

    # Calculate time to expiry
    expiry_date = datetime.strptime(expiry, "%Y-%m-%d")
    days_to_expiry = (expiry_date - datetime.now()).days
    time_to_expiry = max(days_to_expiry / 365.0, 0.001)

    st.subheader(f"{ticker} Options - Expires {expiry} ({days_to_expiry} days)")

    # Fetch options chain
    with st.spinner("Loading options chain..."):
        calls_raw, puts_raw = get_options_chain(ticker, expiry)

    if calls_raw is not None:
        calls_df = process_options_df(calls_raw, current_price, time_to_expiry, "CALL")
        puts_df = process_options_df(puts_raw, current_price, time_to_expiry, "PUT")

        # Display tabs
        tab1, tab2 = st.tabs(["📈 Calls", "📉 Puts"])

        with tab1:
            st.dataframe(
                calls_df.style.format({
                    'Strike': '${:.2f}',
                    'Bid': '${:.2f}',
                    'Ask': '${:.2f}',
                    'Last': '${:.2f}',
                    'Mid': '${:.2f}',
                    'Volume': '{:,.0f}',
                    'Open Int': '{:,.0f}',
                    'Impl Vol': '{:.1%}',
                    'Delta': '{:.2f}',
                }),
                use_container_width=True,
                height=500
            )

        with tab2:
            st.dataframe(
                puts_df.style.format({
                    'Strike': '${:.2f}',
                    'Bid': '${:.2f}',
                    'Ask': '${:.2f}',
                    'Last': '${:.2f}',
                    'Mid': '${:.2f}',
                    'Volume': '{:,.0f}',
                    'Open Int': '{:,.0f}',
                    'Impl Vol': '{:.1%}',
                    'Delta': '{:.2f}',
                }),
                use_container_width=True,
                height=500
            )

        # Download button
        st.divider()
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            excel_file = create_excel_download(
                ticker, calls_df, puts_df, expiry, current_price,
                st.session_state['expirations']
            )
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 Download Excel with P&L Calculator",
                data=excel_file,
                file_name=f"{ticker}_OPTIONS_{expiry}_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        st.caption("💡 Download the Excel file to enter positions and calculate P&L scenarios")

    else:
        st.error("Failed to load options chain")

else:
    st.info("👈 Enter a ticker symbol and click **Load Expirations** to get started")

    st.markdown("""
    ### How to use:
    1. Enter a stock ticker (e.g., CCJ, AAPL, TSLA)
    2. Click **Load Expirations** to see available option dates
    3. Select an expiration from the dropdown
    4. View calls and puts with calculated deltas
    5. Download Excel file for P&L analysis

    ### Excel Features:
    - Enter option positions (+ long, - short)
    - Enter stock positions
    - Set "Stock @ Expiry" for scenario analysis
    - Automatic P&L calculation with breakdown
    """)
