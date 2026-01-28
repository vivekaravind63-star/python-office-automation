import pandas as pd
import os

folder_path = r""   # user changes this (Add your path)
date_column = "Date"                  

all_files = []
first_file = True

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path, file)

        df = pd.read_excel(file_path)

        # Convert date column to short date format
        df[date_column] = pd.to_datetime(df[date_column], errors="coerce") \
                            .dt.strftime("%d-%m-%Y")

        df.dropna(how="all", inplace=True)
        all_files.append(df)

final_df = pd.concat(all_files, ignore_index=True)

final_df.to_excel("Final_Report.xlsx", index=False)

print("Excel files merged Successfully")
