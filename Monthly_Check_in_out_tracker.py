import os
import pandas as pd

parent_folder = input(
    "Enter the path to the parent folder containing subfolders: ").strip()

if not parent_folder or not os.path.isdir(parent_folder):
    print("No valid folder selected.")
    exit()

output_excel = os.path.join(parent_folder, "Monthly Punching Records.xlsx")

# Store all sheets in a dict {sheet_name: DataFrame}
sheets_data = {}

for root, dirs, files in os.walk(parent_folder):
    for file in files:
        if file.endswith(('.xlsx', '.xls')) and "Weekly Punching Records" not in file:
            file_path = os.path.join(root, file)
            try:
                # read everything as text first
                df = pd.read_excel(file_path, dtype=str)

                # ✅ Keep only rows where Att. Date is a valid date
                df["SortDate"] = pd.to_datetime(
                    df["Att. Date"], errors="coerce")
                df = df[df["SortDate"].notna()].copy()

                # Add Day column (use SortDate for reliability)
                df["Day"] = df["SortDate"].dt.day_name()

                # Keep only required columns (original Att. Date untouched!)
                df = df[["Att. Date", "Day", "InTime",
                         "OutTime", "Total Duration", "SortDate"]]

                # Sheet name (max 31 chars)
                sheet_name = os.path.splitext(file)[0][:31]

                # 🔹 Append to dict
                if sheet_name in sheets_data:
                    sheets_data[sheet_name] = pd.concat(
                        [sheets_data[sheet_name], df], ignore_index=True)
                else:
                    sheets_data[sheet_name] = df

            except Exception as e:
                print(f"Error reading {file}: {e}")

# 🔹 Process each sheet: sort by SortDate + add continuous S.No
for sheet_name, data in sheets_data.items():
    data = data.sort_values("SortDate", ascending=True).reset_index(drop=True)
    data.insert(0, "S.No", range(1, len(data) + 1))  # continuous numbering
    data = data.drop(columns=["SortDate"])  # remove helper column
    sheets_data[sheet_name] = data

# 🔹 Finally, write everything to Excel once
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    for sheet_name, data in sheets_data.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Combined Excel written to: {output_excel}")
