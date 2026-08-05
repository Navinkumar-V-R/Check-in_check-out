import os
import pandas as pd

folder_path = input(
    "Enter the path to the folder containing Excel files: ").strip()

if not folder_path or not os.path.isdir(folder_path):
    print("No valid folder selected.")
    exit()

output_excel = os.path.join(folder_path, "Weekly Punching Records.xlsx")
excel_files = [f for f in os.listdir(
    folder_path) if f.endswith(('.xlsx', '.xls'))]

with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)
            # Filter only rows with valid dates in "Att. Date"
            df = df[pd.to_datetime(
                df["Att. Date"], errors="coerce").notna()].copy()
            # Add Day column based on Att. Date
            df["Day"] = pd.to_datetime(df["Att. Date"]).dt.day_name()
            # Keep only required columns
            df = df[["Att. Date", "Day", "InTime",
                    "OutTime", "Total Duration"]]
            # Add S.No column
            df.insert(0, "S.No", range(1, len(df) + 1))
            sheet_name = os.path.splitext(file)[0][:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"Error reading {file}: {e}")

print(f"Combined Excel written to: {output_excel}")
