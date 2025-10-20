import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os

# --- Configuration ---
file_path = r"C:\Users\uditr\Downloads\all_nse1.xlsx+"
base_output_dir = r"C:\Users\uditr\Allproject\stackdata\stockfloder"   # Main output folder

# Read Excel file
df = pd.read_excel(file_path)
print("Columns:", df.columns.tolist())  # optional — to verify column names

# List of intervals
intervals = [ '5m', '15m','30m', '1h', '1d']

# Create base output folder if not exists
os.makedirs(base_output_dir, exist_ok=True)

total_col = 0
list_interval_col_count = []

for interval in intervals:
    # Create a subfolder for each interval

    # for symbols in df.iloc[:, 1]:
    #     print(symbols)
    interval_col_count = 0
    for symbols in df.iloc[:, 1]:
        if pd.isna(symbols):
            print("⚠️ Skipping empty symbol entry.")
            continue
        print(symbols)
        symbol = symbols + ".NS"  # Stock symbol
        # interval = '5m'
        interval_folder = os.path.join(base_output_dir, interval)
        os.makedirs(interval_folder, exist_ok=True)

        # --- Date Range ---

        # end_date = datetime.now()
        # start_date = end_date - timedelta(days=60)
        end_date = datetime.now()

        if interval in ['5m', '15m', '30m']:
            start_date = end_date - timedelta(days=60)
        elif interval == '1h':
            start_date = end_date - timedelta(days=730)  # 2 years
        elif interval == '1d':
            start_date = end_date - timedelta(days=3660)  # 10 years
        else:
            start_date = end_date - timedelta(days=365)




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
        output_file = os.path.join(interval_folder, f"{symbol}_{interval}_data.xlsx")

        try:
            data.to_excel(output_file, index=False)
            print(f"✅ Saved: {output_file}")
            print(f"📊 Total records: {len(data)}")
            total_col += len(data)
            interval_col_count += len(data)
            # print(total_col)
            # print(interval_col_count)
        except PermissionError:
            alt_file = os.path.join(interval_folder, f"{symbol}_{interval}_data{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            data.to_excel(alt_file, index=False)
            print(f"⚠️ File open, saved as: {alt_file}")
        except Exception as e:
            alt_file = os.path.join(interval_folder, f"{symbol}_{interval}_data.csv")
            data.to_csv(alt_file, index=False)
            print(f"⚠️ Excel failed ({e}), saved as CSV: {alt_file}")

    print(f"📊 Total records in {interval} : {interval_col_count}")
    list_interval_col_count.append(interval_col_count)

print(f"Total record in dataset : {total_col}")
j =0
for interval in intervals:
    print(f"📊 Total records in {interval} : {list_interval_col_count[j]}")
    j += 1