import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta

# --- Configuration ---
# Symbol (any stock symbol from Yahoo Finance)
symbol = 'AAPL'
# Interval for the data. Options: 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo
interval = '15m'

# --- Date Range Calculation ---
# yfinance allows fetching intraday data for the last 60 days.
# We will fetch data for the last 30 days as in the original script.
end_date = datetime.now()
start_date = end_date - timedelta(days=30)

print(f"Fetching {interval} data for {symbol} from {start_date.date()} to {end_date.date()}...")

# --- Fetch Data from Yahoo Finance ---
# The yf.download() function returns a pandas DataFrame.
# We set auto_adjust=False to get the 'Adj Close' column, though we don't use it here.
data = yf.download(
    tickers=symbol,
    start=start_date,
    end=end_date,
    interval=interval,
    auto_adjust=False, # Set to False to get Adj Close
    progress=False # Hides the download progress bar
)

# Check if any data was returned
if data.empty:
    print(f"❌ No data found for symbol '{symbol}' in the specified date range.")
else:
    # --- Data Processing ---
    # The index is already a DatetimeIndex representing the start time.
    # We reset the index to turn it into a column.
    data = data.reset_index()

    # Rename the timestamp column to 'TimeStart'
    # The column name can vary ('Datetime' or 'Date'), so we check the first column name
    timestamp_col_name = data.columns[0]
    data = data.rename(columns={timestamp_col_name: 'TimeStart'})

    # Add the 'TimeStop' column by adding the interval duration.
    # We parse the interval string to calculate the timedelta.
    interval_value = int(interval[:-1])
    interval_unit = interval[-1]
    if interval_unit == 'm':
        time_delta = pd.Timedelta(minutes=interval_value)
    elif interval_unit == 'h':
        time_delta = pd.Timedelta(hours=interval_value)
    else:
        # Fallback for daily/weekly intervals, though less common for this script
        time_delta = pd.Timedelta(days=1)

    data['TimeStop'] = data['TimeStart'] + time_delta

    # Reorder columns to match the desired format
    # Note: Yahoo Finance column names are already capitalized ('Open', 'High', etc.)
    data = data[['TimeStart', 'TimeStop', 'Open', 'High', 'Low', 'Close', 'Volume']]


    # --- Save to Excel ---
    output_file = f'{symbol}_{interval}_data_yahoo.xlsx'
    data.to_excel(output_file, index=False)

    print(f"✅ Excel file created successfully: {output_file}")
    print(f"Total records fetched: {len(data)}")


