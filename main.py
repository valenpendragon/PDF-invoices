import pandas as pd
import glob
import os
import pathlib

filepaths = glob.glob("Invoices/*.xlsx")

dataframes = []
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    dataframes.append(df)

print(dataframes)