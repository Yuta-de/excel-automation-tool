import pandas as pd
file_path = r"C:\work\python_study\excel_automation\work1\売上データ元\store_A.xlsx"
df = pd.read_excel(file_path)
print(df)