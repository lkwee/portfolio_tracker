import pandas as pd
import yfinance as yf
from datetime import datetime
import re
pd.options.mode.chained_assignment = None  # default='warn'

TODAY = datetime.now().date()

# Load CSV file
data_path = "path/to/your/transaction_file.xlsx"
df = pd.read_excel(data_path)

df.columns = df.columns.str.strip()
df['NettPrice'] = pd.to_numeric(df['NettPrice'].replace(r'[\$,]', '', regex=True), errors='coerce')
df['Price'] = pd.to_numeric(df['Price'].replace(r'[\$,]', '', regex=True), errors='coerce')
df['date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d', dayfirst=True, errors='coerce')
df['total_shares_held'] = df.groupby('Ticker')['Qty'].cumsum().fillna(0)# Convert 'Ticker' column to string and fill NaN values to avoid errors
df['Ticker'] = df['Ticker'].astype(str).str.strip().fillna('')

# To label ETFs as "ETF" in the Industry column, examples shown in the set below
etf_tickers = {'CSPX.L', 'SPYL.L'}

exchange_rates, industry_info = {}, {}
latest_prices = {}

# For private or in-house funds by Investment Companies, input the names in the exclude_tickers set. 
# Then parse funds_info.txt for excluded tickers' Total Cost and Total Value
exclude_tickers = {"fund1", "fund2"}

funds_info_path = "path/to/your/funds_info.txt"
excluded_cost_value = {}
with open(funds_info_path, 'r') as file:
    lines = file.readlines()
    for i in range(0, len(lines), 2):
        ticker = lines[i].split('=')[0].split()[-1].strip()
        cost = float(lines[i].split('=')[1].strip())
        value = float(lines[i + 1].split('=')[1].strip())
        excluded_cost_value[ticker] = {'Total Cost (SGD)': cost, 'Total Value (SGD)': value}
fallback_start_date = '2000-01-01'

def fetch_prices_and_industry(ticker, start_date):
    """Fetch latest price and industry info for a given ticker."""
    try:
        prices = yf.download(ticker, start=start_date, end=TODAY)['Close']
        if prices.empty:
            print(f"No price data found for {ticker} using yf.download")
            # Trouble shooting if a certain ticker fails to be fetched
            if ticker == "D05.SI":
                latest_prices[ticker] = yf.Ticker('D05.SI').info['regularMarketPrice'] 
                print(f"Fetched price for {ticker} using yf.Ticker.info['regularMarketPrice']: {latest_prices[ticker]}")
            else:
                latest_prices[ticker] = 0
        else:
            latest_prices[ticker] = prices.iloc[-1].values[0] #if not prices.empty else 0
        industry_info[ticker] = yf.Ticker(ticker).info.get('industry', 'Unknown')
    except Exception as e:
        print(f"Failed to fetch data for {ticker}: {e}")
        latest_prices[ticker] = 0
        industry_info[ticker] = 'Unknown'
    return latest_prices, industry_info
    
def fetch_exchange_rate(currency):
    try:
        rate = yf.download(f'{currency}SGD=X', start=fallback_start_date, end=TODAY)['Close']
        # print(type(rate), rate)
        if rate.empty:
            print(f"No data for {currency}. Returning default value 1.")
            return 1
        latest_rate = rate.iloc[-1].values[0]
        # print(f"Latest rate for {currency}: {latest_rate}")
        return latest_rate
    except:
        return 1

def compute_summary(ticker, ticker_data, latest_prices): 
    ticker_data['EffectivePrice'] = ticker_data['NettPrice'].combine_first(ticker_data['Price'])
    total_qty = ticker_data['Qty'].sum()
     
    avg_price = (ticker_data['EffectivePrice'] * ticker_data['Qty']).sum() / total_qty if total_qty > 0 else 0
    latest_price, exchange_rate = latest_prices[ticker], exchange_rates.get(ticker_data['Currency'].iloc[0], 1)
    total_cost, total_value = total_qty * avg_price, total_qty * latest_price
    total_cost_sgd, total_value_sgd, profit_sgd = total_cost * exchange_rate, total_value * exchange_rate, (total_value - total_cost) * exchange_rate
    return {
        'Ticker': ticker, 'Company': ticker_data['Name'].iloc[-1] if 'Name' in ticker_data.columns else None,
        'Industry': 'ETF' if ticker in etf_tickers else industry_info.get(ticker),
        'Currency': ticker_data['Currency'].iloc[0], 'Total Shares': total_qty,
        'Average Purchase Price': avg_price, 'Total Cost': round(total_cost, 2), 'Total Cost (SGD)': round(total_cost_sgd, 2),
        'Latest Price': latest_price, 'Total Value': round(total_value, 2), 'Total Value (SGD)': round(total_value_sgd, 2),
        'Profit (SGD)': round(profit_sgd, 2), 'Exchange Rate': exchange_rate, 'exch_rate_date': TODAY
    }

exchange_rates = {curr: fetch_exchange_rate(curr) for curr in df['Currency'].unique() if curr != 'SGD'}

summary, total_cost_sgd, total_value_sgd, total_profit_sgd = [], 0, 0, 0

# Iterate over unique tickers
for ticker in df['Ticker'].unique():
    if ticker in exclude_tickers or not ticker:
        continue
    ticker_data = df[df['Ticker'] == ticker]
    # Use the earliest date from the data, or the fallback if no valid date is found
    start_date = (
        ticker_data['date'].min().strftime('%Y-%m-%d') 
        if not pd.isna(ticker_data['date'].min()) 
        else fallback_start_date
    )
    latest_prices, industry_info = fetch_prices_and_industry(ticker, start_date) 
    summary_data = compute_summary(ticker, ticker_data, latest_prices)
    if summary_data:
        summary.append(summary_data)
        
for ticker in exclude_tickers:
    if ticker in excluded_cost_value:
        total_cost_sgd = excluded_cost_value[ticker]['Total Cost (SGD)']
        total_value_sgd = excluded_cost_value[ticker]['Total Value (SGD)']
        profit_sgd = total_value_sgd - total_cost_sgd
        if ticker == 'G3B.SI':
            avg_price = 3.207
        else:
            avg_price = 0
            
        summary.append({
            **excluded_cost_value[ticker],
            'Company': ticker,
            'Profit (SGD)': profit_sgd,
            'Total Cost': excluded_cost_value[ticker]['Total Cost (SGD)'],
            'Total Value': excluded_cost_value[ticker]['Total Value (SGD)'],
            'Ticker': ticker,
            'Industry': 'ETF',
            'Currency': 'SGD',
            'Average Purchase Price': avg_price, 
            'Exchange Rate': 1,
            'exch_rate_date': TODAY
        })

# Calculate weightage
summary_df = pd.DataFrame(summary)

# Delete ticker with qty=0 (Shares fully sold)
summary_df = summary_df[~((summary_df['Total Shares']==0) & (~summary_df['Ticker'].isin(exclude_tickers)))]

usd_summary = summary_df[(summary_df['Currency'] == 'USD') & (~summary_df['Ticker'].isin(exclude_tickers))]
non_usd_summary = summary_df[
    (summary_df['Currency'] != 'USD') & 
    (~summary_df['Ticker'].isin(exclude_tickers) | (summary_df['Ticker'] == 'G3B.SI'))
]
usd_weightage = (usd_summary['Total Value (SGD)'] / usd_summary['Total Value (SGD)'].sum() * 100).round(2)
non_usd_weightage = (non_usd_summary['Total Value (SGD)'] / non_usd_summary['Total Value (SGD)'].sum() * 100).round(2)
summary_df.loc[usd_summary.index, 'Weightage (%)'] = usd_weightage
summary_df.loc[non_usd_summary.index, 'Weightage (%)'] = non_usd_weightage

# Calculate the earliest investment date for each ticker
df['Earliest Investment Date'] = df.groupby('Ticker')['date'].transform('min')
# Ensure the column is in the summary DataFrame
summary_df['Earliest Investment Date'] = summary_df['Ticker'].map(
    df.groupby('Ticker')['Earliest Investment Date'].first().dt.strftime('%Y-%m-%d'))

# Parse total amount of dividends that were collected so far
dividends_data = "path/to/your/Dividends.xlsx"
dividends_df = pd.read_excel(dividends_data)

dividends_df.columns = dividends_df.columns.str.strip()
dividends_df['Dividend'] = pd.to_numeric(dividends_df['Dividend'], errors='coerce')
total_dividends_per_ticker = dividends_df.groupby('Ticker')['Dividend'].sum()
summary_df['Dividends'] = summary_df['Ticker'].map(total_dividends_per_ticker).fillna(0)

# Compute grand total
grand_totals = {
    'Ticker': 'Grand Total',
    'Total Cost (SGD)': summary_df['Total Cost (SGD)'].sum(),
    'Total Value (SGD)': summary_df['Total Value (SGD)'].sum(),
    'Profit (SGD)': summary_df['Profit (SGD)'].sum(),
    'Dividends': summary_df['Dividends'].sum()
}
summary_df = pd.concat([summary_df, pd.DataFrame([grand_totals])], ignore_index=True)

print(summary_df)
summary_df.to_csv(f'portfolio_summary_{TODAY}.csv', index=False)