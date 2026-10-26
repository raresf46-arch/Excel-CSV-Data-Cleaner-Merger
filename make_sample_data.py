"""Generate deliberately messy sample files (3 'branches' exporting sales differently)."""
import random
from pathlib import Path
import pandas as pd

random.seed(7)
OUT = Path(__file__).parent / "sample_data"
OUT.mkdir(exist_ok=True)

clients = ["Alpha SRL", "Beta Ltd", "Gamma Co", "Delta SA", "Epsilon Trade", "Zeta Logistics"]
products = ["Sensor Kit", "Cable Set", "Control Unit", "Mounting Kit", "Service Hours"]

def rows(n, month_range):
    out = []
    for i in range(n):
        m = random.choice(month_range)
        d = random.randint(1, 28)
        out.append((pd.Timestamp(2026, m, d), random.choice(clients), random.choice(products),
                    random.randint(1, 20), round(random.uniform(40, 900), 2)))
    return out

# Branch A: clean-ish CSV, ISO dates, dot decimals, but trailing spaces + a duplicate row
a = rows(40, [1, 2, 3])
df_a = pd.DataFrame(a, columns=["Date", "Client", "Product", "Qty", "Unit Price"])
df_a["Date"] = df_a["Date"].dt.strftime("%Y-%m-%d")
df_a["Client"] = df_a["Client"].apply(lambda c: c + "  " if random.random() < 0.3 else c)
df_a = pd.concat([df_a, df_a.iloc[[3, 7]]])  # duplicates
df_a.to_csv(OUT / "branch_A_export.csv", index=False)

# Branch B: semicolon CSV, European formats, different headers, amounts like "1.250,50 RON"
b = rows(35, [2, 3, 4])
df_b = pd.DataFrame(b, columns=["data", "customer", "item", "quantity", "price"])
df_b["data"] = df_b["data"].dt.strftime("%d.%m.%Y")
df_b["customer"] = df_b["customer"].str.upper()
df_b["price"] = df_b["price"].apply(lambda p: f"{p:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " RON")
df_b.loc[5, "quantity"] = None          # missing value
df_b.loc[12, "customer"] = ""           # missing client
df_b.to_csv(OUT / "branch_B_vanzari.csv", index=False, sep=";")

# Branch C: Excel with a title row on top, text dates, blank rows, negative qty typo
c = rows(30, [3, 4, 5])
df_c = pd.DataFrame(c, columns=["Order Date", "Client Name", "Product Name", "Units", "Price/Unit"])
df_c["Order Date"] = df_c["Order Date"].dt.strftime("%d %b %Y")
df_c["Client Name"] = df_c["Client Name"].str.lower()
df_c.loc[8, "Units"] = -4               # data entry error
df_c.loc[15, "Order Date"] = "n/a"      # bad date
blank = pd.DataFrame([[None] * 5], columns=df_c.columns)
df_c = pd.concat([df_c.iloc[:10], blank, df_c.iloc[10:20], blank, df_c.iloc[20:]])
with pd.ExcelWriter(OUT / "branch_C_report_FINAL_v2.xlsx") as w:
    pd.DataFrame([["SALES REPORT - BRANCH C (exported manually)"]]).to_excel(w, index=False, header=False, startrow=0)
    df_c.to_excel(w, index=False, startrow=2)

print("Sample files written to", OUT)
