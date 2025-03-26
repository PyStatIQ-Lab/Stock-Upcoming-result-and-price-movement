import yfinance as yf
import pandas as pd
import os
from openpyxl import load_workbook
import streamlit as st
from datetime import datetime, timedelta

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
        
        latest_close_price = historical_data_1d['Close'].iloc[-1]
        
        # Calculate price trends
        price_1m_ago = historical_data_1m['Close'].iloc[0]
        price_3m_ago = historical_data_3m['Close'].iloc[0]
        
        price_change_1m = ((latest_close_price - price_1m_ago) / price_1m_ago) * 100
        price_change_3m = ((latest_close_price - price_3m_ago) / price_3m_ago) * 100
        
        result['1M Price Change'] = f"{price_change_1m:.2f}%"
        result['3M Price Change'] = f"{price_change_3m:.2f}%"
        
        # Determine price trend
        if price_change_1m > 5 and price_change_3m > 10:
            price_trend = "Strong Uptrend"
        elif price_change_1m > 2 and price_change_3m > 5:
            price_trend = "Moderate Uptrend"
        elif price_change_1m < -5 and price_change_3m < -10:
            price_trend = "Strong Downtrend"
        elif price_change_1m < -2 and price_change_3m < -5:
            price_trend = "Moderate Downtrend"
        else:
            price_trend = "Neutral"
            
        result['Price Trend'] = price_trend
        
    except Exception as e:
        latest_close_price = "N/A"
        result['1M Price Change'] = "N/A"
        result['3M Price Change'] = "N/A"
        result['Price Trend'] = "N/A"

    result['Net Income'] = income_statement.loc['Net Income'] if 'Net Income' in income_statement.index else "N/A"
    result['Operating Income'] = income_statement.loc['Operating Income'] if 'Operating Income' in income_statement.index else \
                                 income_statement.loc['EBIT'] if 'EBIT' in income_statement.index else "N/A"
    
    try:
        eps = income_statement.loc['Earnings Before Interest and Taxes'] / stock.info['sharesOutstanding']
    except KeyError:
        eps = "N/A"
    result['EPS'] = eps
    
    result['Revenue Growth'] = income_statement.loc['Total Revenue'].pct_change().iloc[-1] if 'Total Revenue' in income_statement.index else "N/A"
    
    result['Retained Earnings'] = balance_sheet.loc['Retained Earnings'] if 'Retained Earnings' in balance_sheet.index else "N/A"
    result['Cash Reserves'] = balance_sheet.loc['Cash'] if 'Cash' in balance_sheet.index else "N/A"
    
    try:
        result['Debt-to-Equity Ratio'] = balance_sheet.loc['Total Debt'] / balance_sheet.loc['Stockholders Equity'] if 'Total Debt' in balance_sheet.index and 'Stockholders Equity' in balance_sheet.index else "N/A"
    except KeyError:
        result['Debt-to-Equity Ratio'] = "N/A"
    
    result['Working Capital'] = balance_sheet.loc['Total Assets'] - balance_sheet.loc['Total Liabilities Net Minority Interest'] if 'Total Assets' in balance_sheet.index and 'Total Liabilities Net Minority Interest' in balance_sheet.index else "N/A"
    
    result['Dividend Payout Ratio'] = stock.info.get('dividendYield', "N/A")
    result['Dividend Yield'] = result['Dividend Payout Ratio']
    
    result['Free Cash Flow'] = cash_flow.loc['Free Cash Flow'] if 'Free Cash Flow' in cash_flow.index else "N/A"
    
    if not dividends.empty:
        result['Dividend Growth Rate'] = dividends.pct_change().mean()
    else:
        result['Dividend Growth Rate'] = "N/A"
    
    result['Latest Close Price'] = latest_close_price
    result['Dividend Percentage'] = "N/A"
    
    if not dividends.empty:
        predicted_dividend_amount = dividends.iloc[-1]
        if latest_close_price != "N/A":
            dividend_percentage = (predicted_dividend_amount / latest_close_price) * 100
            result['Dividend Percentage'] = dividend_percentage
        
        past_dividends = dividends.tail(10)
        result['Past Dividends'] = past_dividends.tolist()
        
        date_diffs = past_dividends.index.to_series().diff().dropna()
        if not date_diffs.empty:
            avg_diff = date_diffs.mean()
            last_dividend_date = past_dividends.index[-1]
            next_dividend_date = last_dividend_date + avg_diff
            result['Next Dividend Date'] = str(next_dividend_date)
        else:
            result['Next Dividend Date'] = 'N/A'

        result['Predicted Dividend Amount'] = predicted_dividend_amount
    else:
        result['Next Dividend Date'] = 'N/A'
        result['Predicted Dividend Amount'] = 'N/A'
        result['Dividend Percentage'] = "N/A"

    # Add upcoming earnings date and expectation
    try:
        # Get earnings calendar
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
            if days_until_earnings <= 14:  # Earnings within 2 weeks
                if "Uptrend" in price_trend:
                    result['Earnings Expectation'] = "Positive (Price rising before earnings)"
                elif "Downtrend" in price_trend:
                    result['Earnings Expectation'] = "Negative (Price falling before earnings)"
                else:
                    result['Earnings Expectation'] = "Neutral (No clear trend)"
            else:
                result['Earnings Expectation'] = "Too early to predict (Earnings >2 weeks away)"
        else:
            result['Next Earnings Date'] = "N/A"
            result['Days Until Earnings'] = "N/A"
            result['Earnings Expectation'] = "N/A"
    except Exception as e:
        result['Next Earnings Date'] = "N/A"
        result['Days Until Earnings'] = "N/A"
        result['Earnings Expectation'] = "N/A"

    return result

# Rest of your code remains the same...
# [Keep all the existing functions and Streamlit app code below]
