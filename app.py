import yfinance as yf
import pandas as pd
import os
from openpyxl import load_workbook
import streamlit as st
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import numpy as np

# Path to the stocks.xlsx file
STOCKS_FILE_PATH = 'stocks.xlsx'

# Function to fetch earnings trend data (FIXED: Added proper error handling)
def get_earnings_trend(ticker):
    stock = yf.Ticker(ticker)
    result = {}
    
    try:
        # Get earnings history (FIXED: Added check for empty data)
        earnings = stock.earnings
        if earnings is not None and not earnings.empty and 'Estimate' in earnings.columns and 'Actual' in earnings.columns:
            earnings['Surprise (%)'] = (earnings['Actual'] - earnings['Estimate']) / earnings['Estimate'] * 100
            result['Past Earnings'] = earnings
            result['Avg Surprise (%)'] = earnings['Surprise (%)'].mean()
            result['Beat Rate'] = len(earnings[earnings['Surprise (%)'] > 0]) / len(earnings)
        else:
            result['Past Earnings'] = None
            result['Avg Surprise (%)'] = np.nan
            result['Beat Rate'] = np.nan

        # Get earnings calendar (FIXED: Added column existence checks)
        calendar = stock.calendar
        if calendar is not None and not calendar.empty:
            result['Next Earnings Date'] = calendar.index[0].strftime('%Y-%m-%d') if not calendar.empty else "N/A"
            
            if 'Earnings Estimate' in calendar.columns:
                result['Next EPS Estimate'] = calendar.iloc[0]['Earnings Estimate']
            else:
                result['Next EPS Estimate'] = "N/A"
                
            if 'Revenue Estimate' in calendar.columns:
                result['Next Revenue Estimate'] = calendar.iloc[0]['Revenue Estimate']
            else:
                result['Next Revenue Estimate'] = "N/A"
        else:
            result['Next Earnings Date'] = "N/A"
            result['Next EPS Estimate'] = "N/A"
            result['Next Revenue Estimate'] = "N/A"

    except Exception as e:
        st.error(f"Error fetching earnings data for {ticker}: {str(e)}")
        result['Past Earnings'] = None
        result['Avg Surprise (%)'] = np.nan
        result['Beat Rate'] = np.nan
        result['Next Earnings Date'] = "N/A"
        result['Next EPS Estimate'] = "N/A"
        result['Next Revenue Estimate'] = "N/A"
    
    return result

# Function to predict earnings direction (FIXED: Added null checks)
def predict_earnings(ticker, price_trend="Neutral"):
    earnings_data = get_earnings_trend(ticker)
    prediction = "N/A"
    confidence = "N/A"
    
    try:
        if not np.isnan(earnings_data['Avg Surprise (%)']) and not np.isnan(earnings_data['Beat Rate']):
            avg_surprise = earnings_data['Avg Surprise (%)']
            beat_rate = earnings_data['Beat Rate']
            
            # Prediction logic (FIXED: Added more nuanced conditions)
            if avg_surprise > 5 and beat_rate > 0.7:
                prediction = "Likely BEAT (Strong historical performance)"
                confidence = "High"
            elif avg_surprise > 2 and beat_rate > 0.6:
                prediction = "Likely BEAT (Good historical performance)"
                confidence = "Medium-High"
            elif avg_surprise < -5 and beat_rate < 0.3:
                prediction = "Likely MISS (Poor historical performance)"
                confidence = "High"
            elif avg_surprise < -2 and beat_rate < 0.4:
                prediction = "Likely MISS (Weak historical performance)"
                confidence = "Medium-High"
            else:
                prediction = "Likely MEET (In-line with estimates)"
                confidence = "Medium"
            
            # Adjust based on price trend (FIXED: Added proper string checks)
            if price_trend and isinstance(price_trend, str):
                if "Uptrend" in price_trend and confidence in ["Medium", "Medium-High"]:
                    prediction += " + Positive price momentum"
                    confidence = "Medium-High"
                elif "Downtrend" in price_trend and confidence in ["Medium", "Medium-High"]:
                    prediction += " - Negative price momentum"
                    confidence = "Medium-High"
    
    except Exception as e:
        st.error(f"Prediction error for {ticker}: {str(e)}")
    
    return {
        'Prediction': prediction,
        'Confidence': confidence,
        'Next EPS Estimate': earnings_data['Next EPS Estimate'],
        'Next Revenue Estimate': earnings_data['Next Revenue Estimate'],
        'Avg Surprise (%)': f"{earnings_data['Avg Surprise (%)']:.1f}%" if not np.isnan(earnings_data['Avg Surprise (%)']) else "N/A",
        'Beat Rate': f"{earnings_data['Beat Rate']*100:.1f}%" if not np.isnan(earnings_data['Beat Rate']) else "N/A"
    }

# Function to fetch all financial data (FIXED: Multiple error handling improvements)
def get_financial_data(ticker):
    stock = yf.Ticker(ticker)
    result = {'Ticker': ticker}
    
    try:
        # Get financial statements with proper error handling
        income_statement = stock.financials
        balance_sheet = stock.balance_sheet
        cash_flow = stock.cashflow
        dividends = stock.dividends
        info = stock.info
    except Exception as e:
        st.error(f"Error fetching financial data for {ticker}: {str(e)}")
        return None

    # Price data and trends (FIXED: Added empty data checks)
    try:
        historical_data = stock.history(period="3mo")
        if not historical_data.empty and len(historical_data) >= 60:  # Ensure enough data points
            latest_close = historical_data['Close'].iloc[-1]
            sma_50 = historical_data['Close'].rolling(20).mean().iloc[-1]
            sma_200 = historical_data['Close'].rolling(60).mean().iloc[-1]
            
            result['Latest Price'] = f"${latest_close:.2f}"
            result['50-day SMA'] = f"${sma_50:.2f}"
            result['200-day SMA'] = f"${sma_200:.2f}"
            
            # Price trend analysis (FIXED: Added index checks)
            if len(historical_data) >= 20:
                price_change_1m = (latest_close - historical_data['Close'].iloc[-20]) / historical_data['Close'].iloc[-20] * 100
                price_change_3m = (latest_close - historical_data['Close'].iloc[0]) / historical_data['Close'].iloc[0] * 100
                
                if price_change_1m > 5 and price_change_3m > 10:
                    price_trend = "Strong Uptrend"
                elif price_change_1m > 2 and price_change_3m > 5:
                    price_trend = "Moderate Uptrend"
                elif price_change_1m < -5 and price_change_3m < -10:
                    price_trend = "Strong Downtrend"
                elif price_change_1m < -2 and price_change_3m < -5:
                    price_trend = "Moderate Downtrend"
                else:
                    price_trend = "Neutral Trend"
                
                result['Price Trend'] = price_trend
                result['Historical Data'] = historical_data
        else:
            result['Latest Price'] = "N/A"
            result['Price Trend'] = "N/A"
            result['Historical Data'] = None
    except Exception as e:
        st.warning(f"Price data error for {ticker}: {str(e)}")
        result['Latest Price'] = "N/A"
        result['Price Trend'] = "N/A"
        result['Historical Data'] = None

    # Financial metrics (FIXED: Added multiple safety checks)
    try:
        if income_statement is not None and 'Net Income' in income_statement.index and 'sharesOutstanding' in info:
            eps = income_statement.loc['Net Income'].iloc[0]/info['sharesOutstanding']
            result['EPS'] = f"${eps:.2f}"
        else:
            result['EPS'] = "N/A"
            
        if income_statement is not None and 'Total Revenue' in income_statement.index:
            revenue_growth = income_statement.loc['Total Revenue'].pct_change().iloc[0]*100 if len(income_statement.loc['Total Revenue']) > 1 else 0
            result['Revenue Growth'] = f"{revenue_growth:.1f}%"
        else:
            result['Revenue Growth'] = "N/A"
            
        if (income_statement is not None and 'Net Income' in income_statement.index 
            and 'Total Revenue' in income_statement.index and income_statement.loc['Total Revenue'].iloc[0] != 0):
            profit_margin = income_statement.loc['Net Income'].iloc[0]/income_statement.loc['Total Revenue'].iloc[0]*100
            result['Profit Margin'] = f"{profit_margin:.1f}%"
        else:
            result['Profit Margin'] = "N/A"
            
        if 'trailingPE' in info:
            result['P/E Ratio'] = f"{info['trailingPE']:.1f}"
        else:
            result['P/E Ratio'] = "N/A"
    except Exception as e:
        st.warning(f"Financial metric error for {ticker}: {str(e)}")
        result['EPS'] = "N/A"
        result['Revenue Growth'] = "N/A"
        result['Profit Margin'] = "N/A"
        result['P/E Ratio'] = "N/A"

    # Add earnings prediction (FIXED: Added fallback for missing price trend)
    price_trend = result.get('Price Trend', "Neutral")
    earnings_pred = predict_earnings(ticker, price_trend)
    result.update(earnings_pred)

    return result

# Function to save results to Excel (FIXED: Improved file handling)
def save_to_excel(results, filename="stock_analysis_results.xlsx"):
    try:
        # Prepare data for DataFrame
        data_for_excel = []
        for result in results:
            row = {
                'Ticker': result['Ticker'],
                'Latest Price': result['Latest Price'],
                '50-day SMA': result['50-day SMA'],
                '200-day SMA': result['200-day SMA'],
                'Price Trend': result['Price Trend'],
                'EPS': result['EPS'],
                'P/E Ratio': result['P/E Ratio'],
                'Revenue Growth': result['Revenue Growth'],
                'Profit Margin': result['Profit Margin'],
                'Earnings Prediction': result['Prediction'],
                'Confidence': result['Confidence'],
                'Next EPS Estimate': result['Next EPS Estimate'],
                'Next Revenue Estimate': result['Next Revenue Estimate'],
                'Avg Surprise (%)': result['Avg Surprise (%)'],
                'Beat Rate': result['Beat Rate']
            }
            data_for_excel.append(row)
        
        results_df = pd.DataFrame(data_for_excel)
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        
        # Save to Excel with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_with_ts = f"{os.path.splitext(filename)[0]}_{timestamp}.xlsx"
        
        results_df.to_excel(filename_with_ts, index=False)
        st.success(f"Results saved to {filename_with_ts}")
        return filename_with_ts
    except Exception as e:
        st.error(f"Error saving to Excel: {str(e)}")
        return None

# Streamlit App Configuration
st.set_page_config(
    page_title="Stock Analysis Pro", 
    layout="wide",
    page_icon="📈"
)

# Custom CSS (FIXED: Improved styling)
st.markdown("""
    <style>
        .header { 
            text-align: center; 
            margin-bottom: 30px;
            padding: 20px;
            background: linear-gradient(135deg, #6e8efb, #a777e3);
            color: white;
            border-radius: 10px;
        }
        .metric-card { 
            padding: 15px; 
            border-radius: 10px; 
            background-color: #f8f9fa; 
            margin-bottom: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            border-left: 4px solid #4e73df;
        }
        .positive { color: #2ecc71; font-weight: bold; }
        .negative { color: #e74c3c; font-weight: bold; }
        .neutral { color: #f39c12; font-weight: bold; }
        .prediction-card { 
            background-color: #e8f4fc; 
            border-left: 4px solid #3498db;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 10px;
        }
        .stButton>button {
            background-color: #4e73df;
            color: white;
            border-radius: 5px;
            padding: 8px 16px;
            border: none;
        }
        .stButton>button:hover {
            background-color: #2e59d9;
        }
    </style>
""", unsafe_allow_html=True)

# Header (FIXED: Improved layout)
st.markdown("""
    <div class="header">
        <h1>📈 Stock Earnings Predictor Pro</h1>
        <p>Advanced earnings prediction using financial metrics and technical analysis</p>
    </div>
""", unsafe_allow_html=True)

# Main App Logic (FIXED: Added try-except blocks)
try:
    if os.path.exists(STOCKS_FILE_PATH):
        try:
            symbols_df = pd.read_excel(STOCKS_FILE_PATH)
            
            if 'Symbol' not in symbols_df.columns:
                st.error("Error: The file must contain a 'Symbol' column with stock tickers")
            else:
                stock_options = symbols_df['Symbol'].unique().tolist()
                selected_stocks = st.multiselect(
                    "Select Stocks to Analyze", 
                    stock_options,
                    help="Choose one or more stocks for analysis"
                )
                
                if st.button('Analyze Selected Stocks', key='analyze_btn') and selected_stocks:
                    with st.spinner('Analyzing stocks...'):
                        all_results = []
                        progress_bar = st.progress(0)
                        
                        for i, ticker in enumerate(selected_stocks):
                            progress_bar.progress((i+1)/len(selected_stocks))
                            result = get_financial_data(ticker)
                            if result:
                                all_results.append(result)
                        
                        progress_bar.empty()
                        
                        if not all_results:
                            st.warning("No results to display")
                        else:
                            for result in all_results:
                                st.markdown(f"## {result['Ticker']} Analysis")
                                
                                # Layout columns
                                col1, col2 = st.columns([1, 2])
                                
                                with col1:
                                    # Earnings Prediction Card
                                    st.markdown(f"""
                                        <div class="prediction-card">
                                            <h3>🎯 Earnings Prediction</h3>
                                            <p><b>Prediction:</b> <span class="{'positive' if 'BEAT' in result['Prediction'] else 'negative' if 'MISS' in result['Prediction'] else 'neutral'}">{result['Prediction']}</span></p>
                                            <p><b>Confidence:</b> {result['Confidence']}</p>
                                            <p><b>Next EPS Estimate:</b> {result['Next EPS Estimate']}</p>
                                            <p><b>Next Revenue Estimate:</b> {result['Next Revenue Estimate']}</p>
                                            <p><b>Avg Surprise (%):</b> {result['Avg Surprise (%)']}</p>
                                            <p><b>Beat Rate:</b> {result['Beat Rate']}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                                    
                                    # Price Trends
                                    st.markdown("### 📊 Price Trends")
                                    st.markdown(f"""
                                        <div class="metric-card">
                                            <p><b>Latest Price:</b> {result['Latest Price']}</p>
                                            <p><b>50-day SMA:</b> {result['50-day SMA']}</p>
                                            <p><b>200-day SMA:</b> {result['200-day SMA']}</p>
                                            <p><b>Trend:</b> <span class="{'positive' if 'Uptrend' in result['Price Trend'] else 'negative' if 'Downtrend' in result['Price Trend'] else 'neutral'}">{result['Price Trend']}</span></p>
                                        </div>
                                    """, unsafe_allow_html=True)
                                
                                with col2:
                                    # Financial Metrics
                                    st.markdown("### 💰 Financial Metrics")
                                    st.markdown(f"""
                                        <div class="metric-card">
                                            <div style="column-count: 2;">
                                                <p><b>EPS:</b> {result['EPS']}</p>
                                                <p><b>P/E Ratio:</b> {result['P/E Ratio']}</p>
                                                <p><b>Revenue Growth:</b> {result['Revenue Growth']}</p>
                                                <p><b>Profit Margin:</b> {result['Profit Margin']}</p>
                                            </div>
                                        </div>
                                    """, unsafe_allow_html=True)
                                    
                                    # Price Chart
                                    if result['Historical Data'] is not None:
                                        st.markdown("### 📈 Price Chart (3 Months)")
                                        fig, ax = plt.subplots(figsize=(10, 4))
                                        ax.plot(result['Historical Data'].index, 
                                               result['Historical Data']['Close'], 
                                               label='Price', color='#4e73df')
                                        ax.set_title(f"{result['Ticker']} Price Movement")
                                        ax.set_xlabel("Date")
                                        ax.set_ylabel("Price ($)")
                                        ax.grid(True, linestyle='--', alpha=0.7)
                                        st.pyplot(fig)
                                    else:
                                        st.warning("No historical price data available")
                                
                                st.markdown("---")
                            
                            # Export button
                            if st.button('📤 Export Results to Excel'):
                                saved_file = save_to_excel(all_results)
                                if saved_file:
                                    st.success(f"Successfully saved to {saved_file}")
                                    with open(saved_file, "rb") as file:
                                        st.download_button(
                                            label="Download Excel File",
                                            data=file,
                                            file_name=os.path.basename(saved_file),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
        except Exception as e:
            st.error(f"Error reading stock file: {str(e)}")
    else:
        st.error(f"File not found: {STOCKS_FILE_PATH}. Please ensure the file exists in the correct location.")
except Exception as e:
    st.error(f"An unexpected error occurred: {str(e)}")

# Footer
st.markdown("""
    <div style="text-align: center; margin-top: 50px; color: #7f8c8d; font-size: 0.9em;">
        <p>Stock Analysis Pro | Powered by yFinance and Streamlit</p>
        <p>Data may be delayed. Use for informational purposes only.</p>
    </div>
""", unsafe_allow_html=True)
