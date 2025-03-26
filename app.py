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

# Function to fetch earnings trend data
def get_earnings_trend(ticker):
    stock = yf.Ticker(ticker)
    result = {}
    
    try:
        # Get earnings history
        earnings = stock.earnings
        if earnings is not None and not earnings.empty:
            earnings['Surprise (%)'] = (earnings['Actual'] - earnings['Estimate']) / earnings['Estimate'] * 100
            result['Past Earnings'] = earnings
            result['Avg Surprise (%)'] = earnings['Surprise (%)'].mean()
            result['Beat Rate'] = len(earnings[earnings['Surprise (%)'] > 0]) / len(earnings)
        else:
            result['Past Earnings'] = None
            result['Avg Surprise (%)'] = "N/A"
            result['Beat Rate'] = "N/A"

        # Get earnings calendar
        calendar = stock.calendar
        if calendar is not None and not calendar.empty:
            result['Next Earnings Date'] = calendar.index[0].strftime('%Y-%m-%d')
            if 'Earnings Estimate' in calendar.columns:
                result['Next EPS Estimate'] = calendar.iloc[0]['Earnings Estimate']
            if 'Revenue Estimate' in calendar.columns:
                result['Next Revenue Estimate'] = calendar.iloc[0]['Revenue Estimate']
        else:
            result['Next Earnings Date'] = "N/A"
            result['Next EPS Estimate'] = "N/A"
            result['Next Revenue Estimate'] = "N/A"

    except Exception as e:
        st.error(f"Error fetching earnings data for {ticker}: {e}")
        result['Past Earnings'] = None
        result['Avg Surprise (%)'] = "N/A"
        result['Beat Rate'] = "N/A"
        result['Next Earnings Date'] = "N/A"
        result['Next EPS Estimate'] = "N/A"
        result['Next Revenue Estimate'] = "N/A"
    
    return result

# Function to predict earnings direction
def predict_earnings(ticker, price_trend):
    earnings_data = get_earnings_trend(ticker)
    prediction = "N/A"
    confidence = "N/A"
    
    if earnings_data['Avg Surprise (%)'] != "N/A":
        avg_surprise = earnings_data['Avg Surprise (%)']
        beat_rate = earnings_data['Beat Rate']
        
        # Prediction logic
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
        
        # Adjust based on price trend
        if "Uptrend" in price_trend and confidence in ["Medium", "Medium-High"]:
            prediction += " + Positive price momentum"
            confidence = "Medium-High"
        elif "Downtrend" in price_trend and confidence in ["Medium", "Medium-High"]:
            prediction += " - Negative price momentum"
            confidence = "Medium-High"
    
    return {
        'Prediction': prediction,
        'Confidence': confidence,
        'Next EPS Estimate': earnings_data['Next EPS Estimate'],
        'Next Revenue Estimate': earnings_data['Next Revenue Estimate'],
        'Avg Surprise (%)': earnings_data['Avg Surprise (%)'],
        'Beat Rate': f"{earnings_data['Beat Rate']*100:.1f}%" if earnings_data['Beat Rate'] != "N/A" else "N/A"
    }

# Function to fetch all financial data (updated)
def get_financial_data(ticker):
    stock = yf.Ticker(ticker)
    result = {'Ticker': ticker}
    
    try:
        # Get financial statements
        income_statement = stock.financials
        balance_sheet = stock.balance_sheet
        cash_flow = stock.cashflow
        dividends = stock.dividends
        info = stock.info
    except Exception as e:
        st.error(f"Error fetching financial data for {ticker}: {e}")
        return None

    # Price data and trends
    try:
        historical_data = stock.history(period="3mo")
        if not historical_data.empty:
            latest_close = historical_data['Close'].iloc[-1]
            sma_50 = historical_data['Close'].rolling(20).mean().iloc[-1]
            sma_200 = historical_data['Close'].rolling(60).mean().iloc[-1]
            
            result['Latest Price'] = f"${latest_close:.2f}"
            result['50-day SMA'] = f"${sma_50:.2f}"
            result['200-day SMA'] = f"${sma_200:.2f}"
            
            # Price trend analysis
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
        st.warning(f"Price data error for {ticker}: {e}")
        result['Latest Price'] = "N/A"
        result['Price Trend'] = "N/A"
        result['Historical Data'] = None

    # Financial metrics
    try:
        result['EPS'] = f"${income_statement.loc['Net Income'].iloc[0]/info['sharesOutstanding']:.2f}" if 'Net Income' in income_statement.index else "N/A"
        result['Revenue Growth'] = f"{income_statement.loc['Total Revenue'].pct_change().iloc[0]*100:.1f}%" if 'Total Revenue' in income_statement.index else "N/A"
        result['Profit Margin'] = f"{income_statement.loc['Net Income'].iloc[0]/income_statement.loc['Total Revenue'].iloc[0]*100:.1f}%" if 'Net Income' in income_statement.index and 'Total Revenue' in income_statement.index else "N/A"
        result['P/E Ratio'] = f"{info['trailingPE']:.1f}" if 'trailingPE' in info else "N/A"
    except:
        result['EPS'] = "N/A"
        result['Revenue Growth'] = "N/A"
        result['Profit Margin'] = "N/A"
        result['P/E Ratio'] = "N/A"

    # Add earnings prediction
    if result['Price Trend'] != "N/A":
        earnings_pred = predict_earnings(ticker, result['Price Trend'])
        result.update(earnings_pred)
    else:
        result.update({
            'Prediction': "N/A",
            'Confidence': "N/A",
            'Next EPS Estimate': "N/A",
            'Next Revenue Estimate': "N/A",
            'Avg Surprise (%)': "N/A",
            'Beat Rate': "N/A"
        })

    return result

# Rest of your existing functions (save_to_excel, plot_stock_performance) remain the same
# [Previous code for these functions goes here]

# Streamlit App
st.set_page_config(page_title="Stock Analysis Pro", layout="wide")

# Custom CSS
st.markdown("""
    <style>
        .header { text-align: center; margin-bottom: 30px; }
        .metric-card { 
            padding: 15px; border-radius: 10px; 
            background-color: #f0f2f6; margin-bottom: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .positive { color: #2ecc71; font-weight: bold; }
        .negative { color: #e74c3c; font-weight: bold; }
        .neutral { color: #f39c12; font-weight: bold; }
        .prediction-card { 
            background-color: #e8f4fc; 
            border-left: 4px solid #3498db;
            padding: 15px;
            margin-bottom: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="header">
        <h1>📈 Stock Earnings Predictor</h1>
        <p>Predict upcoming earnings results using financial metrics and price trends</p>
    </div>
""", unsafe_allow_html=True)

# Main App
if os.path.exists(STOCKS_FILE_PATH):
    symbols_df = pd.read_excel(STOCKS_FILE_PATH)
    
    if 'Symbol' not in symbols_df.columns:
        st.error("Error: The file must contain a 'Symbol' column")
    else:
        stock_options = symbols_df['Symbol'].tolist()
        selected_stocks = st.multiselect("Select Stocks", stock_options, help="Choose stocks to analyze")
        
        if st.button('Analyze Stocks') and selected_stocks:
            all_results = []
            progress_bar = st.progress(0)
            
            for i, ticker in enumerate(selected_stocks):
                progress_bar.progress((i+1)/len(selected_stocks))
                result = get_financial_data(ticker)
                if result:
                    all_results.append(result)
            
            progress_bar.empty()
            
            for result in all_results:
                st.markdown(f"## {result['Ticker']} Analysis")
                
                # Layout columns
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    # Earnings Prediction Card
                    st.markdown("""
                        <div class="prediction-card">
                            <h3>🎯 Earnings Prediction</h3>
                            <p><b>Prediction:</b> <span class="{'positive' if 'BEAT' in result['Prediction'] else 'negative' if 'MISS' in result['Prediction'] else 'neutral'}">{}</span></p>
                            <p><b>Confidence:</b> {}</p>
                            <p><b>Next EPS Estimate:</b> {}</p>
                            <p><b>Next Revenue Estimate:</b> {}</p>
                            <p><b>Avg Surprise (%):</b> {}</p>
                            <p><b>Beat Rate:</b> {}</p>
                        </div>
                    """.format(
                        result['Prediction'],
                        result['Confidence'],
                        result['Next EPS Estimate'] if result['Next EPS Estimate'] != "N/A" else "N/A",
                        result['Next Revenue Estimate'] if result['Next Revenue Estimate'] != "N/A" else "N/A",
                        f"{result['Avg Surprise (%)']:.1f}%" if result['Avg Surprise (%)'] != "N/A" else "N/A",
                        result['Beat Rate']
                    ), unsafe_allow_html=True)
                    
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
                        st.markdown("### 📈 Price Chart")
                        fig, ax = plt.subplots(figsize=(10, 4))
                        ax.plot(result['Historical Data'].index, result['Historical Data']['Close'], label='Price')
                        ax.set_title(f"{result['Ticker']} 3-Month Price")
                        ax.grid(True)
                        st.pyplot(fig)
                
                st.markdown("---")
            
            # Export button
            if st.button('📤 Export Results to Excel'):
                save_to_excel(all_results, "stock_analysis_results.xlsx")
else:
    st.error(f"File not found: {STOCKS_FILE_PATH}")

# Footer
st.markdown("""
    <div style="text-align: center; margin-top: 50px; color: #7f8c8d;">
        <p>Stock Analysis Pro | Powered by yFinance</p>
    </div>
""", unsafe_allow_html=True)
