import pandas as pd
import os
from glob import glob

# get source directory
source_dir = r"C:\Users\7075685\Documents\#PRJ\2025017_iproケアパックツール修正\002_ケアパック集計\原本"

# set output file name
output_file ="output.xlsx"

# set destination directory
destination_dir = r"C:\Users\7075685\Documents\#PRJ\2025017_iproケアパックツール修正\002_ケアパック集計"

# set source sheet name
source_sheet_name = ["オンサイト", "センドバック", "Nパッケージ"]  # Example sheet names

# set merged column range(A:AQ)
column_range = list(range(0, 43))  # Adjust as needed

# set array of combined sheet name
combine_sheets = {sheet: [] for sheet in source_sheet_name}

# get target excel files(.xlsx)
files = glob(os.path.join(source_dir, "*.xlsx"))

for file in files:
    for sheet in source_sheet_name:
        try:
            df = pd.read_excel(file, sheet_name=sheet, usecols=column_range)
            df["元ファイル"] = os.path.basename(file)
            combine_sheets[sheet].append(df)
        except Exception as e:
            print(f"エラー: {os.path.basename(file)} の {sheet} シート → {e}")

# save merged data to a new excel file
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet, df_list in combine_sheets.items():
        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            combined_df.to_excel(writer, sheet_name=sheet, index=False)

# finish program msg
print (f"merged successfully.: {output_file}")
