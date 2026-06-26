import pandas as pd

excel_file = "data/Engine Programme Management.xlsx"

df = pd.read_excel(
    excel_file,
    sheet_name="radb"
)

print("Columns:")
for column in df.columns:
    print(f"- {column}")

service_limit = 2500
warning_limit = 2250
🟢 Verde     si km < 2250
🟡 Amarillo  si 2250 <= km < 2500
🔴 Rojo      si km >= 2500

keywords = ["boost", "p21", "carter", "cam", "knock"]

for keyword in keywords:
    print(f"\nColumnas que contienen '{keyword}':")
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            print("-", col)
            