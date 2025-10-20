import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os

# --- Configuration ---
file_path = r"C:\Users\uditr\Downloads\Average MCAP_July2024ToDecember 2024 (1).xlsx"
base_output_dir = r"C:\Users\uditr\Allproject\stackdata\.venv\Scripts\python.exe C:\Users\uditr\Allproject\stackdata\stockfloder"   # Main output folder

# Read Excel file
df = pd.read_excel(file_path)
print("Columns:", df.columns.tolist())  # optional — to verify column names

# List of intervals
intervals = ['1m', '5m', '15m', '1h', '1d']

# Create base output folder if not exists
os.makedirs(base_output_dir, exist_ok=True)

for interval in intervals:
    # Create a subfolder for each interval
    interval_folder = os.path.join(base_output_dir, interval)
    os.makedirs(interval_folder, exist_ok=True)

    # (You can loop over all symbols — currently using one for demo)
    symbol = "HDFCBANK" + ".NS"

    # --- Date Range ---
    end_date = datetime.now()
    start_date = end_date - timedelta(days=10000)

    print(f"\nFetching {interval} data for {symbol} from {start_date.date()} to {end_date.date()}...")

    # --- Fetch data ---
    data = yf.download(
        tickers=symbol,
        start=start_date,
        end=end_date,
        interval=interval,
        auto_adjust=False,
        progress=False
    )

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.droplevel(1)

    if data.empty:
        print(f"❌ No data found for symbol '{symbol}' at interval {interval}.")
        continue

    # --- Prepare dataframe ---
    data = data.reset_index()
    data = data.rename(columns={data.columns[0]: 'TimeStart'})

    # Convert timezone-aware datetime to naive
    if isinstance(data['TimeStart'].dtype, pd.DatetimeTZDtype):
        data['TimeStart'] = data['TimeStart'].dt.tz_localize(None)

    # Add TimeStop column
    interval_value = int(interval[:-1])
    interval_unit = interval[-1]

    if interval_unit == 'm':
        delta = pd.Timedelta(minutes=interval_value)
    elif interval_unit == 'h':
        delta = pd.Timedelta(hours=interval_value)
    else:
        delta = pd.Timedelta(days=1)

    data['TimeStop'] = data['TimeStart'] + delta
    data = data[['TimeStart', 'TimeStop', 'Open', 'High', 'Low', 'Close', 'Volume']]

    # --- Save file in its interval folder ---
    output_file = os.path.join(interval_folder, f"{symbol}_{interval}_data_yahoo.xlsx")

    try:
        data.to_excel(output_file, index=False)
        print(f"✅ Saved: {output_file}")
        print(f"📊 Total records: {len(data)}")
    except PermissionError:
        alt_file = os.path.join(interval_folder, f"{symbol}_{interval}_data_yahoo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        data.to_excel(alt_file, index=False)
        print(f"⚠️ File open, saved as: {alt_file}")
    except Exception as e:
        alt_file = os.path.join(interval_folder, f"{symbol}_{interval}_data_yahoo.csv")
        data.to_csv(alt_file, index=False)
        print(f"⚠️ Excel failed ({e}), saved as CSV: {alt_file}")
