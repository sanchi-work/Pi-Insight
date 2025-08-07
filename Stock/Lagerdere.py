import os
import pandas as pd
import re

# === CONFIGURATION ===
parent_dir = r"C:\Users\sanch\OneDrive\Desktop\Lag"
sheet_name = "Stocks"

month_mapping = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
    "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
}
reverse_month_mapping = {v: k for k, v in month_mapping.items()}

def load_excel_safely(filepath, sheet=None):
    try:
        return pd.read_excel(filepath, sheet_name=sheet or 0)
    except Exception as e:
        print(f"❌ Failed to load {filepath}: {e}")
        return pd.DataFrame()

def process_single_month(year: int, month_str: str):
    year_folder = f"Lag{year}"
    folder_path = os.path.join(parent_dir, year_folder)

    if not os.path.exists(folder_path):
        print(f"🚫 Folder does not exist: {folder_path}")
        return pd.DataFrame()

    month_num = reverse_month_mapping.get(month_str)
    if not month_num:
        print(f"❌ Invalid month: {month_str}")
        return pd.DataFrame()

    file_prefix = month_num
    month_files = [f for f in os.listdir(folder_path) if f.startswith(file_prefix) and f.endswith(".xlsx") and not f.startswith("~$")]

    if not month_files:
        print(f"🚫 No files found for {month_str} {year}")
        return pd.DataFrame()

    all_data = []

    for filename in month_files:
        filepath = os.path.join(folder_path, filename)
        print(f"> Reading: {filename}")
        df = load_excel_safely(filepath, sheet=sheet_name)
        if df.empty:
            continue

        required_cols = ['Material - EAN/UPC (Key)', 'Material - Text']
        if not all(col in df.columns for col in required_cols):
            print(f"❌ Missing required columns in {filename}. Skipping.")
            continue

        df['Material - EAN/UPC (Key)'] = df['Material - EAN/UPC (Key)'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        df['Material - Text'] = df['Material - Text'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        df['New_MaterialText'] = df['Material - EAN/UPC (Key)'] + "#" + df['Material - Text']

        df['Month'] = month_str
        df['Year'] = year
        df['Source_File'] = filename

        all_data.append(df)

    if not all_data:
        print(f"🚫 No valid data for {month_str} {year}")
        return pd.DataFrame()

    return pd.concat(all_data, ignore_index=True)

if __name__ == "__main__":
    # === Process only July 2025 ===
    year = 2025
    month_str = "Jul"

    df_july_2025 = process_single_month(year, month_str)

    if not df_july_2025.empty:
        output_path = os.path.join(parent_dir, f"Lag{year}", f"Lagardere_Stocks_July_{year}_Processed.xlsx")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_july_2025.to_excel(writer, sheet_name="July", index=False)
        print(f"✅ July {year} output saved: {output_path}")
    else:
        print("⚠️ July file is missing or invalid.")
