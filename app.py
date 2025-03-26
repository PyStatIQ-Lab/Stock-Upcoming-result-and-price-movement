import yfinance as yf
import pandas as pd
import os
from openpyxl import load_workbook
import streamlit as st
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

# Path to the stocks.xlsx file
STOCKS_FILE_PATH = 'stocks.xlsx'  # Change this to the correct path if needed

# Function to fetch data for a given stock ticker
def get_financial_data(ticker):
    stock = yf.Ticker(ticker)
    result = {'Ticker': ticker}
    
    try:
        income_statement = stock.financials
        balance_sheet = stock.balance_sheet
        cash_flow = stock.cashflow
        dividends = stock.dividends
        info = stock.info
    except Exception as e:
        st.error(f"Error fetching financial data for {ticker}: {e}")
        return None

    try:
        # Get historical data for different time periods
        historical_data_1m = stock.history(period="1mo")
        historical_data_3m = stock.history(period="3mo")
        historical_data_1d = stock.history(period="1d")
        historical_data_1y = stock.history(period="1y")
        
        latest_close_price = historical_data_1d['Close'].iloc[-1] if not historical_data_1d.empty else None
        
        # Calculate price trends
        price_1m_ago = historical_data_1m['Close'].iloc[0] if not historical_data_1m.empty else None
        price_3m_ago = historical_data_3m['Close'].iloc[0] if not historical_data_3m.empty else None
        price_1y_ago = historical_data_1y['Close'].iloc[0] if not historical_data_1y.empty else None
        
        price_change_1m = ((latest_close_price - price_1m_ago) / price_1m_ago) * 100 if price_1m_ago and latest_close_price else None
        price_change_3m = ((latest_close_price - price_3m_ago) / price_3m_ago) * 100 if price_3m_ago and latest_close_price else None
        price_change_1y = ((latest_close_price - price_1y_ago) / price_1y_ago) * 100 if price_1y_ago and latest_close_price else None
        
        result['1M Price Change'] = f"{price_change_1m:.2f}%" if price_change_1m is not None else "N/A"
        result['3M Price Change'] = f"{price_change_3m:.2f}%" if price_change_3m is not None else "N/A"
        result['1Y Price Change'] = f"{price_change_1y:.2f}%" if price_change_1y is not None else "N/A"
        
        # Determine price trend
        price_trend = "N/A"
        if price_change_1m is not None and price_change_3m is not None:
            if price_change_1m > 10 and price_change_3m > 20:
                price_trend = "Very Strong Uptrend"
            elif price_change_1m > 5 and price_change_3m > 10:
                price_trend = "Strong Uptrend"
            elif price_change_1m > 2 and price_change_3m > 5:
                price_trend = "Moderate Uptrend"
            elif price_change_1m < -10 and price_change_3m < -20:
                price_trend = "Very Strong Downtrend"
            elif price_change_1m < -5 and price_change_3m < -10:
                price_trend = "Strong Downtrend"
            elif price_change_1m < -2 and price_change_3m < -5:
                price_trend = "Moderate Downtrend"
            else:
                price_trend = "Neutral Trend"
                
        result['Price Trend'] = price_trend
        
        # Store historical data for visualization
        result['Historical Data'] = historical_data_1y if not historical_data_1y.empty else None
        
    except Exception as e:
        st.warning(f"Could not fetch complete price data for {ticker}: {e}")
        latest_close_price = None
        result['1M Price Change'] = "N/A"
        result['3M Price Change'] = "N/A"
        result['1Y Price Change'] = "N/A"
        result['Price Trend'] = "N/A"
        result['Historical Data'] = None

    # Basic financial metrics
    try:
        result['Net Income'] = income_statement.loc['Net Income'].iloc[0] if 'Net Income' in income_statement.index else "N/A"
    except:
        result['Net Income'] = "N/A"
        
    try:
        result['Operating Income'] = income_statement.loc['Operating Income'].iloc[0] if 'Operating Income' in income_statement.index else \
                                   income_statement.loc['EBIT'].iloc[0] if 'EBIT' in income_statement.index else "N/A"
    except:
        result['Operating Income'] = "N/A"
    
    try:
        shares_outstanding = info['sharesOutstanding']
        eps = income_statement.loc['Net Income'].iloc[0] / shares_outstanding
        result['EPS'] = f"${eps:.2f}"
    except:
        result['EPS'] = "N/A"
    
    try:
        revenue_growth = income_statement.loc['Total Revenue'].pct_change().iloc[-1] * 100
        result['Revenue Growth'] = f"{revenue_growth:.2f}%"
    except:
        result['Revenue Growth'] = "N/A"
    
    try:
        result['Retained Earnings'] = balance_sheet.loc['Retained Earnings'].iloc[0] if 'Retained Earnings' in balance_sheet.index else "N/A"
    except:
        result['Retained Earnings'] = "N/A"
        
    try:
        result['Cash Reserves'] = balance_sheet.loc['Cash'].iloc[0] if 'Cash' in balance_sheet.index else "N/A"
    except:
        result['Cash Reserves'] = "N/A"
    
    try:
        debt = balance_sheet.loc['Total Debt'].iloc[0]
        equity = balance_sheet.loc['Stockholders Equity'].iloc[0]
        result['Debt-to-Equity Ratio'] = f"{debt/equity:.2f}"
    except:
        result['Debt-to-Equity Ratio'] = "N/A"
    
    try:
        assets = balance_sheet.loc['Total Assets'].iloc[0]
        liabilities = balance_sheet.loc['Total Liabilities Net Minority Interest'].iloc[0]
        result['Working Capital'] = f"${assets - liabilities:,.2f}"
    except:
        result['Working Capital'] = "N/A"
    
    try:
        result['Dividend Yield'] = f"{info.get('dividendYield', 0) * 100:.2f}%" if 'dividendYield' in info else "N/A"
    except:
        result['Dividend Yield'] = "N/A"
    
    try:
        result['Free Cash Flow'] = cash_flow.loc['Free Cash Flow'].iloc[0] if 'Free Cash Flow' in cash_flow.index else "N/A"
    except:
        result['Free Cash Flow'] = "N/A"
    
    # Dividend information
    if not dividends.empty:
        try:
            result['Dividend Growth Rate'] = f"{dividends.pct_change().mean() * 100:.2f}%"
        except:
            result['Dividend Growth Rate'] = "N/A"
    else:
        result['Dividend Growth Rate'] = "N/A"
    
    result['Latest Close Price'] = f"${latest_close_price:.2f}" if latest_close_price is not None else "N/A"
    
    if not dividends.empty:
        try:
            predicted_dividend_amount = dividends.iloc[-1]
            if latest_close_price is not None:
                dividend_percentage = (predicted_dividend_amount / float(latest_close_price)) * 100
                result['Dividend Percentage'] = f"{dividend_percentage:.2f}%"
            
            past_dividends = dividends.tail(4)  # Last 4 dividends
            result['Past Dividends'] = [f"${x:.2f}" for x in past_dividends.tolist()]
            
            date_diffs = past_dividends.index.to_series().diff().dropna()
            if not date_diffs.empty:
                avg_diff = date_diffs.mean()
                last_dividend_date = past_dividends.index[-1]
                next_dividend_date = last_dividend_date + avg_diff
                result['Next Dividend Date'] = str(next_dividend_date.date())
                result['Days Until Dividend'] = (next_dividend_date.date() - datetime.now().date()).days
            else:
                result['Next Dividend Date'] = 'N/A'
                result['Days Until Dividend'] = 'N/A'

            result['Predicted Dividend Amount'] = f"${predicted_dividend_amount:.2f}"
        except:
            result['Next Dividend Date'] = 'N/A'
            result['Days Until Dividend'] = 'N/A'
            result['Predicted Dividend Amount'] = 'N/A'
            result['Dividend Percentage'] = "N/A"
            result['Past Dividends'] = []
    else:
        result['Next Dividend Date'] = 'N/A'
        result['Days Until Dividend'] = 'N/A'
        result['Predicted Dividend Amount'] = 'N/A'
        result['Dividend Percentage'] = "N/A"
        result['Past Dividends'] = []

    # Earnings information
    try:
        earnings_dates = stock.calendar
        if earnings_dates is not None and not earnings_dates.empty:
            next_earnings = earnings_dates.iloc[0]
            result['Next Earnings Date'] = str(next_earnings.name.date())
            
            # Calculate days until earnings
            today = datetime.now().date()
            earnings_date = next_earnings.name.date()
            days_until_earnings = (earnings_date - today).days
            
            result['Days Until Earnings'] = days_until_earnings
            
            # Determine expectation based on price trend and days until earnings
            if days_until_earnings <= 7:  # Earnings within 1 week
                if "Uptrend" in price_trend:
                    result['Earnings Expectation'] = "Very Positive (Strong uptrend right before earnings)"
                    result['Earnings Confidence'] = "High"
                elif "Downtrend" in price_trend:
                    result['Earnings Expectation'] = "Very Negative (Strong downtrend right before earnings)"
                    result['Earnings Confidence'] = "High"
                else:
                    result['Earnings Expectation'] = "Neutral (No clear trend before earnings)"
                    result['Earnings Confidence'] = "Medium"
            elif days_until_earnings <= 14:  # Earnings within 2 weeks
                if "Uptrend" in price_trend:
                    result['Earnings Expectation'] = "Positive (Uptrend building before earnings)"
                    result['Earnings Confidence'] = "Medium-High"
                elif "Downtrend" in price_trend:
                    result['Earnings Expectation'] = "Negative (Downtrend building before earnings)"
                    result['Earnings Confidence'] = "Medium-High"
                else:
                    result['Earnings Expectation'] = "Neutral (No clear trend yet)"
                    result['Earnings Confidence'] = "Medium"
            else:
                result['Earnings Expectation'] = "Too early to predict (Earnings >2 weeks away)"
                result['Earnings Confidence'] = "Low"
        else:
            result['Next Earnings Date'] = "N/A"
            result['Days Until Earnings'] = "N/A"
            result['Earnings Expectation'] = "N/A"
            result['Earnings Confidence'] = "N/A"
    except Exception as e:
        result['Next Earnings Date'] = "N/A"
        result['Days Until Earnings'] = "N/A"
        result['Earnings Expectation'] = "N/A"
        result['Earnings Confidence'] = "N/A"

    return result

# Function to save results to an Excel file
def save_to_excel(results, filename="dividend_predictions.xlsx"):
    try:
        # Prepare data for DataFrame
        data_for_excel = []
        for result in results:
            row = {
                'Ticker': result['Ticker'],
                'Latest Price': result['Latest Close Price'],
                '1M Change': result['1M Price Change'],
                '3M Change': result['3M Price Change'],
                '1Y Change': result['1Y Price Change'],
                'Price Trend': result['Price Trend'],
                'Net Income': result['Net Income'],
                'EPS': result['EPS'],
                'Revenue Growth': result['Revenue Growth'],
                'Debt-to-Equity': result['Debt-to-Equity Ratio'],
                'Dividend Yield': result['Dividend Yield'],
                'Next Dividend Date': result['Next Dividend Date'],
                'Predicted Dividend': result['Predicted Dividend Amount'],
                'Next Earnings Date': result['Next Earnings Date'],
                'Days Until Earnings': result['Days Until Earnings'],
                'Earnings Expectation': result['Earnings Expectation'],
                'Earnings Confidence': result['Earnings Confidence']
            }
            data_for_excel.append(row)
        
        results_df = pd.DataFrame(data_for_excel)
        
        if os.path.exists(filename):
            book = load_workbook(filename)
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                writer.book = book
                results_df.to_excel(writer, index=False, sheet_name='New Results')
            st.success(f"Results appended to {filename}")
        else:
            results_df.to_excel(filename, index=False)
            st.success(f"Results saved to new file {filename}")
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")

# Function to plot stock performance
def plot_stock_performance(ticker, historical_data):
    if historical_data is None or historical_data.empty:
        return None
    
    plt.figure(figsize=(10, 5))
    plt.plot(historical_data.index, historical_data['Close'], label='Closing Price')
    plt.title(f'{ticker} 1-Year Performance')
    plt.xlabel('Date')
    plt.ylabel('Price ($)')
    plt.grid(True)
    plt.legend()
    return plt

# Streamlit App
st.set_page_config(page_title="Stock Dividend Predictions", layout="wide")

# Custom CSS
st.markdown("""
    <style>
        .header-logo {
            display: block;
            margin-left: auto;
            margin-right: auto;
            width: 25%;
        }
        .metric-card {
            padding: 15px;
            border-radius: 10px;
            background-color: #0d46b8;
            margin-bottom: 10px;
        }
        .positive {
            color: green;
            font-weight: bold;
        }
        .negative {
            color: red;
            font-weight: bold;
        }
        .neutral {
            color: orange;
            font-weight: bold;
        }
        /* Hide GitHub icons and fork button */
        .css-1v0mbdj, .css-1b22hs3, footer, .css-1r6ntm8 { 
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# Display Header Logo
st.markdown('<img class="header-logo" src="https://pystatiq.com/images/pystatIQ_logo.png" alt="Header Logo">', unsafe_allow_html=True)

st.title('Stock Dividend Prediction and Financial Analysis')

# Read the stock symbols from the local stocks.xlsx file
if os.path.exists(STOCKS_FILE_PATH):
    symbols_df = pd.read_excel(STOCKS_FILE_PATH)

    # Check if the 'Symbol' column exists
    if 'Symbol' not in symbols_df.columns:
        st.error("The file must contain a 'Symbol' column with stock tickers.")
    else:
        # Let the user select stocks from the file
        stock_options = symbols_df['Symbol'].tolist()
        selected_stocks = st.multiselect("Select Stock Symbols", stock_options, help="Choose one or more stocks to analyze")

        # Button to start the data fetching process
        if st.button('Fetch Financial Data') and selected_stocks:
            all_results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, ticker in enumerate(selected_stocks):
                status_text.text(f"Processing {ticker} ({i+1}/{len(selected_stocks)})...")
                progress_bar.progress((i + 1) / len(selected_stocks))
                
                result = get_financial_data(ticker)
                if result is not None:
                    all_results.append(result)
            
            progress_bar.empty()
            status_text.empty()
            
            if all_results:
                st.success("Analysis complete!")
                
                # Display results for each stock
                for result in all_results:
                    st.markdown(f"## {result['Ticker']} Analysis")
                    
                    # Create columns for layout
                    col1, col2 = st.columns([2, 3])
                    
                    with col1:
                        # Key metrics
                        st.markdown("### Key Metrics")
                        
                        # Price information
                        st.markdown(f"""
                            <div class="metric-card">
                                <b>Latest Price:</b> {result['Latest Close Price']}<br>
                                <b>1M Change:</b> <span class="{'positive' if '%' in result['1M Price Change'] and float(result['1M Price Change'].replace('%','')) > 0 else 'negative' if '%' in result['1M Price Change'] and float(result['1M Price Change'].replace('%','')) < 0 else ''}">{result['1M Price Change']}</span><br>
                                <b>3M Change:</b> <span class="{'positive' if '%' in result['3M Price Change'] and float(result['3M Price Change'].replace('%','')) > 0 else 'negative' if '%' in result['3M Price Change'] and float(result['3M Price Change'].replace('%','')) < 0 else ''}">{result['3M Price Change']}</span><br>
                                <b>1Y Change:</b> <span class="{'positive' if '%' in result['1Y Price Change'] and float(result['1Y Price Change'].replace('%','')) > 0 else 'negative' if '%' in result['1Y Price Change'] and float(result['1Y Price Change'].replace('%','')) < 0 else ''}">{result['1Y Price Change']}</span><br>
                                <b>Price Trend:</b> <span class="{'positive' if 'Uptrend' in result['Price Trend'] else 'negative' if 'Downtrend' in result['Price Trend'] else 'neutral'}">{result['Price Trend']}</span>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        # Dividend information
                        st.markdown("### Dividend Information")
                        st.markdown(f"""
                            <div class="metric-card">
                                <b>Dividend Yield:</b> {result['Dividend Yield']}<br>
                                <b>Next Dividend Date:</b> {result['Next Dividend Date']}<br>
                                <b>Days Until Dividend:</b> {result['Days Until Dividend']}<br>
                                <b>Predicted Amount:</b> {result['Predicted Dividend Amount']}<br>
                                <b>Dividend Growth Rate:</b> {result['Dividend Growth Rate']}<br>
                                <b>Past Dividends:</b> {', '.join(result['Past Dividends']) if result['Past Dividends'] else 'N/A'}
                            </div>
                        """, unsafe_allow_html=True)
                        
                        # Earnings information
                        st.markdown("### Earnings Information")
                        st.markdown(f"""
                            <div class="metric-card">
                                <b>Next Earnings Date:</b> {result['Next Earnings Date']}<br>
                                <b>Days Until Earnings:</b> {result['Days Until Earnings']}<br>
                                <b>Expectation:</b> <span class="{'positive' if 'Positive' in result['Earnings Expectation'] else 'negative' if 'Negative' in result['Earnings Expectation'] else 'neutral'}">{result['Earnings Expectation']}</span><br>
                                <b>Confidence:</b> {result['Earnings Confidence']}
                            </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        # Financial metrics
                        st.markdown("### Financial Metrics")
                        st.markdown(f"""
                            <div class="metric-card">
                                <b>EPS:</b> {result['EPS']}<br>
                                <b>Net Income:</b> {result['Net Income']}<br>
                                <b>Operating Income:</b> {result['Operating Income']}<br>
                                <b>Revenue Growth:</b> {result['Revenue Growth']}<br>
                                <b>Free Cash Flow:</b> {result['Free Cash Flow']}<br>
                                <b>Cash Reserves:</b> {result['Cash Reserves']}<br>
                                <b>Retained Earnings:</b> {result['Retained Earnings']}<br>
                                <b>Working Capital:</b> {result['Working Capital']}<br>
                                <b>Debt-to-Equity:</b> {result['Debt-to-Equity Ratio']}
                            </div>
                        """, unsafe_allow_html=True)
                        
                        # Price chart
                        st.markdown("### Price Performance")
                        fig = plot_stock_performance(result['Ticker'], result['Historical Data'])
                        if fig:
                            st.pyplot(fig)
                        else:
                            st.warning("No historical price data available for chart")
                    
                    st.markdown("---")
                
                # Button to save the results to Excel
                if st.button('Save Results to Excel'):
                    save_to_excel(all_results)
            else:
                st.warning("No results to display")

else:
    st.error(f"{STOCKS_FILE_PATH} not found. Please ensure the file exists.")

# Footer
st.markdown("""
    <div style="text-align: center; margin-top: 50px;">
        <p><strong>App Code:</strong> Stock-Dividend-Prediction</p>
        <p>For support, please email: <a href="mailto:support@pystatiq.com">support@pystatiq.com</a></p>
        <img src="https://predictram.com/images/logo.png" width="100" style="margin-top: 20px;">
    </div>
""", unsafe_allow_html=True)
