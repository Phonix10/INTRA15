import pandas as pd
from tiingo import TiingoClient
from datetime import datetime, timedelta

# Tiingo API key configuration
config = {
    'api_key': '717d713804ee57c4c0d4340edb6e9f0b34f40fe8'
}

client = TiingoClient(config)

# Time range: only last 30 days for intraday data
end_date = datetime.now()
start_date = end_date - timedelta(days=30)

# Symbol (US stock only)
symbol = 'AAPL'
print(f"Fetching 15-min data for {symbol} from {start_date.date()} to {end_date.date()}...")

# Fetch 15-min interval data
data = client.get_dataframe(
    symbol,
    startDate=start_date.strftime('%Y-%m-%d'),
    endDate=end_date.strftime('%Y-%m-%d') ,
    frequency='15Min'
)

# Rename columns to clean format
data = data.rename(columns={
    'open': 'Open',
    'high': 'High',
    'low': 'Low',
    'close': 'Close',
    'adjClose': 'AdjClose',
    'volume': 'Volume'
})

# Add TimeStart and TimeStop columns
data['TimeStart'] = data.index
data['TimeStop'] = data.index + pd.Timedelta(minutes=15)

# Reorder
data = data[['TimeStart', 'TimeStop', 'Open', 'High', 'Low', 'Close']]

# Save to Excel
output_file = f'{symbol}_15min_data.xlsx'
data.to_excel(output_file, index=False)

print(f"✅ Excel file created successfully: {output_file}")
print(f"Total records fetched: {len(data)}")


