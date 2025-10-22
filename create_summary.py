import pandas as pd
import os

# --- Configuration ---
# IMPORTANT: Make sure these paths match the ones in your original script.

# 1. Path to your original Excel file with the list of stock symbols.
input_symbols_file = r"C:\Users\uditr\Downloads\all_nse1.xlsx"

# 2. Path to the main folder where all the stock data subfolders (5m, 15m, etc.) are saved.
base_data_dir = r"C:\Users\uditr\Allproject\stackdata\stockfloder"

# 3. Name and location for the new summary Excel file that will be created.
output_summary_file = r"C:\Users\uditr\Allproject\stackdata\stock_data_summary.xlsx"

# --- Main Script ---

# List of intervals that were downloaded
intervals = ['5m', '15m', '30m', '1h', '1d']

# Step 1: Read the original list of stocks.
try:
    df_symbols = pd.read_excel(input_symbols_file)
    # Assumes the first column is the stock name and the second is the symbol.
    df_symbols.rename(columns={df_symbols.columns[0]: 'Name', df_symbols.columns[1]: 'Symbol'}, inplace=True)
    print(f"✅ Successfully read {len(df_symbols)} symbols from '{os.path.basename(input_symbols_file)}'.")
except FileNotFoundError:
    print(f"❌ ERROR: The input file was not found at '{input_symbols_file}'. Please check the path.")
    exit()
except Exception as e:
    print(f"❌ An error occurred while reading the symbols file: {e}")
    exit()

# Step 2: Prepare a list to hold the summary data for each stock.
summary_data = []

print("\n🔍 Analyzing downloaded data to create summary...")
print("-" * 50)

# Step 3: Loop through each stock in your list.
for index, row in df_symbols.iterrows():
    symbol = row['Symbol']
    name = row['Name']

    if pd.isna(symbol):
        print(f"⚠️ Skipping empty symbol entry for name: '{name}'.")
        continue

    print(f"Processing: {symbol} ({name})")

    # This dictionary will store the final counts for the current stock.
    stock_summary = {'Stock Symbol': symbol, 'Stock Name': name}

    # Step 4: For each stock, check each interval folder for its data file.
    for interval in intervals:
        column_name = f'{interval} Rows'
        row_count = 0  # Default row count is 0.

        # Construct the expected path to the downloaded Excel file.
        file_name = f"{symbol}.NS_{interval}_data.xlsx"
        file_path = os.path.join(base_data_dir, interval, file_name)

        # Check if the file exists.
        if os.path.exists(file_path):
            try:
                # If it exists, read it and count the rows.
                data_df = pd.read_excel(file_path)
                row_count = len(data_df)
            except Exception as e:
                print(f"  - Could not read Excel file for {interval}. Error: {e}")
                # As a fallback, check for a .csv file as created in your original script.
                csv_file_name = f"{symbol}.NS_{interval}_data.csv"
                csv_file_path = os.path.join(base_data_dir, interval, csv_file_name)
                if os.path.exists(csv_file_path):
                    try:
                        data_df = pd.read_csv(csv_file_path)
                        row_count = len(data_df)
                        print(f"  - Found and read CSV fallback instead for {interval}.")
                    except Exception as csv_e:
                        print(f"  - Could not read CSV fallback. Error: {csv_e}")

        # Add the final count to our summary dictionary for this stock.
        stock_summary[column_name] = row_count

    # Add the completed dictionary for the stock to our main list.
    summary_data.append(stock_summary)

# Step 5: Convert the list of summaries into a pandas DataFrame.
summary_df = pd.DataFrame(summary_data)

# Step 6: Save the final DataFrame to the new Excel file.
try:
    summary_df.to_excel(output_summary_file, index=False)
    print("-" * 50)
    print(f"\n🎉 Success! Summary file created at: {output_summary_file}")
except Exception as e:
    print(f"\n❌ ERROR: Failed to save the summary file. Reason: {e}")