import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os

# --- Configuration ---

file_path = r"C:\Users\uditr\Downloads\Average MCAP_July2024ToDecember 2024 (1).xlsx"

df = pd.read_excel(file_path)

# Assuming the column with symbols is named "Symbol" or second column — let's check
print("Columns:", df.columns.tolist())  # optional — to verify column names

# If the symbol column is the 2nd column (index 1)
l = ['1m', '5m', '15m', '1h', '1d']
for i in l:
    # for symbols in df.iloc[:, 1]:
    # print(symbols)
    symbol =  "HDFCBANK" + ".NS"           # Stock symbol
    interval = i            # Interval (1m, 5m, 15m, 1h, 1d, etc.)

    # --- Date Range ---
    end_date = datetime.now()
    start_date = end_date - timedelta(days=730)

    print(f"Fetching {interval} data for {symbol} from {start_date.date()} to {end_date.date()}...")

    # --- Fetch data ---
    data = yf.download(
        tickers=symbol,
        start=start_date,
        end=end_date,
        interval=interval,
        auto_adjust=False,
        progress=False
    )

    # --- Fix MultiIndex columns (if any) ---
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.droplevel(1)

    # --- Check if data was returned ---
    if data.empty:
        print(f"❌ No data found for symbol '{symbol}' in the specified date range.")
    else:
        data = data.reset_index()

        # Rename timestamp column
        timestamp_col = data.columns[0]
        data = data.rename(columns={timestamp_col: 'TimeStart'})

        # ✅ Convert timezone-aware datetime to timezone-naive
        if isinstance(data['TimeStart'].dtype, pd.DatetimeTZDtype):
            data['TimeStart'] = data['TimeStart'].dt.tz_localize(None)

        # Add 'TimeStop' column based on interval
        interval_value = int(interval[:-1])
        interval_unit = interval[-1]

        if interval_unit == 'm':
            delta = pd.Timedelta(minutes=interval_value)
        elif interval_unit == 'h':
            delta = pd.Timedelta(hours=interval_value)
        else:
            delta = pd.Timedelta(days=1)

        data['TimeStop'] = data['TimeStart'] + delta

        # Reorder columns
        data = data[['TimeStart', 'TimeStop', 'Open', 'High', 'Low', 'Close', 'Volume']]

        # --- Save to Excel ---
        output_file = f"{symbol}_{interval}_data_yahoo.xlsx"

        try:
            import openpyxl
        except ImportError:
            print("⚠️ Installing openpyxl...")
            os.system("pip install openpyxl")

        try:
            data.to_excel(output_file, index=False)
            print(f"✅ Excel file created successfully: {output_file}")
            print(f"Total records fetched: {len(data)}")

        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            alt_file = f"{symbol}_{interval}_data_yahoo_{timestamp}.xlsx"
            data.to_excel(alt_file, index=False)
            print(f"⚠️ File was open. Saved as: {alt_file}")
            print(f"Total records fetched: {len(data)}")

        except Exception as e:
            # Fallback to CSV if Excel fails for any other reason
            alt_file = f"{symbol}_{interval}_data_yahoo.csv"
            data.to_csv(alt_file, index=False)
            print(f"⚠️ Excel write failed ({e}). Saved as CSV: {alt_file}")
            print(f"Total records fetched: {len(data)}")
