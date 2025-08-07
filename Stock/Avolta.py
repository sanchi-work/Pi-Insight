import os
import pandas as pd

base_path = r"C:\Users\sanch\OneDrive - Pi Insight & Research Ltd\Pi Insight\Automations_Sanchi\Github\Stock\Avolta"
all_data = []

# List folders like Avo_24, Avo_25
for year_folder in os.listdir(base_path):
    year_path = os.path.join(base_path, year_folder)

    if os.path.isdir(year_path) and year_folder.startswith("Avo_"):
        year = year_folder.split("Avo_")[-1].strip()

        # Inside each year folder, list month folders like 01_Jan
        for month_folder in os.listdir(year_path):
            month_path = os.path.join(year_path, month_folder)

            if os.path.isdir(month_path) and len(month_folder) >= 2 and month_folder[:2].isdigit():
                month = month_folder[:2]

                # Inside each month folder, look for .xlsx files
                for file in os.listdir(month_path):
                    if file.lower().endswith(".xlsx"):
                        file_path = os.path.join(month_path, file)
                        try:
                            df = pd.read_excel(file_path)
                            df.columns = [col.strip().lower() for col in df.columns]
                            df['Year'] = year
                            df['Month'] = month
                            all_data.append(df)
                        except Exception as e:
                            print(f"Error reading {file_path}: {e}")

# Combine all dataframes
if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
    print("Combined Data Preview:")
    print(combined_df.head())
else:
    print("No data found.")
# Export to CSV
output_file = os.path.join(base_path, "combined_stock_data.csv")
combined_df.to_csv(output_file, index=False)
print(f"Exported combined data to: {output_file}")
