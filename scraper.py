import os
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --------------------
# CONFIG
# --------------------
BASE_URL = "https://www.gw2bltc.com/en/tp/search"
PARAMS = {
    "profit-min": 500,
    "profit-pct-min": 0,
    "profit-pct-max": 100,
    "sold-day-min": 0,
    "bought-day-min": 0,
    "ipg": 200,
    "sort": "profit-pct",
    "page": 1
}

OVERCUT_PCT_DEFAULT = 1.10
UNDERCUT_PCT_DEFAULT = 0.90
QTY_DEFAULT = 1

scrape_time_dt = datetime.now()
scrape_time_str = scrape_time_dt.strftime("%Y-%m-%d %H:%M")
safe_title = scrape_time_str.replace(":", "-")

output_file = "scraper-results.xlsx"

# --------------------
# HELPERS
# --------------------
def parse_gold_silver(td):
    gold = silver = 0
    for span in td.find_all("span"):
        classes = span.get("class", [])
        if "cur-t1c" in classes:  # gold
            try:
                gold = int(span.get_text(strip=True))
            except ValueError:
                gold = 0
        elif "cur-t1b" in classes:  # silver
            try:
                silver = int(span.get_text(strip=True))
            except ValueError:
                silver = 0
    return round(gold + silver / 100, 2)

def parse_int(td):
    txt = td.get_text(strip=True).replace(",", "")
    try:
        return int(txt)
    except ValueError:
        return 0

# --------------------
# SCRAPE
# --------------------
all_rows = []
while True:
    print(f"Fetching page {PARAMS['page']}...")
    try:
        r = requests.get(BASE_URL, params=PARAMS, timeout=20)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        break

    soup = BeautifulSoup(r.text, "html.parser")
    rows = soup.select("table.table-result tr")[1:]
    if not rows:
        print("No more rows found — stopping scrape.")
        break

    for row in rows:
        cols = row.find_all("td")
        if len(cols) < 12:
            continue

        item_name = cols[1].get_text(strip=True)
        link_tag = cols[1].find("a", href=True)
        item_link = f"https://www.gw2bltc.com{link_tag['href']}" if link_tag else ""
        buy = parse_gold_silver(cols[3])
        sell = parse_gold_silver(cols[2])
        demand = parse_int(cols[7])
        supply = parse_int(cols[6])
        bought = parse_int(cols[10])
        sold = parse_int(cols[8])
        bids = parse_int(cols[11])
        offers = parse_int(cols[9])

        all_rows.append([
            item_name, item_link, scrape_time_str, buy, sell,
            demand, supply, bought, sold, bids, offers
        ])

    PARAMS["page"] += 1

if not all_rows:
    print("No data scraped.")
    exit()

# --------------------
# NEW SCRAPE DF
# --------------------
df_new = pd.DataFrame(all_rows, columns=[
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers"
])
df_new["Overcut (%)"] = OVERCUT_PCT_DEFAULT
df_new["Undercut (%)"] = UNDERCUT_PCT_DEFAULT
df_new["Overcut (g)"] = 0
df_new["Undercut (g)"] = 0
df_new["Qty"] = QTY_DEFAULT
df_new["Theoretical Profit"] = 0
df_new["Amount Received"] = 0
df_new["ROI (%)"] = 0
df_new["Demand-Supply Gap (%)"] = 0
df_new["Buy Order Placed"] = False
df_new["Sell Order Placed"] = False
df_new["Sold (manual)"] = False

final_column_order = [
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers",
    "Overcut (%)", "Undercut (%)", "Overcut (g)", "Undercut (g)",
    "Qty", "Theoretical Profit", "Amount Received", "ROI (%)",
    "Demand-Supply Gap (%)", "Buy Order Placed", "Sell Order Placed", "Sold (manual)"
]
df_new = df_new[final_column_order]

# --------------------
# MERGE WITH EXISTING FILE
# --------------------
if os.path.exists(output_file):
    df_existing = pd.read_excel(output_file)
    if "Buy Order Placed" in df_existing.columns:
        df_existing = df_existing[df_existing["Buy Order Placed"] == True]
    else:
        df_existing = pd.DataFrame(columns=final_column_order)
else:
    df_existing = pd.DataFrame(columns=final_column_order)

# Ensure old rows have only values (no formulas)
df_existing = df_existing.fillna("")

# Combine
df_combined = pd.concat([df_existing, df_new], ignore_index=True)
df_combined.to_excel(output_file, index=False)

# --------------------
# FORMAT EXCEL
# --------------------
wb = load_workbook(output_file)
ws = wb.active
ws.title = safe_title
ws.freeze_panes = 'B2'

max_row = ws.max_row
header_to_idx = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], start=1)}
def L(name): return get_column_letter(header_to_idx[name])

# Apply formats & formulas only to NEW rows
start_new = len(df_existing) + 2  # +2 because Excel rows start at 1 and row 1 is header

for row in range(start_new, max_row + 1):
    ws[f'{L("Buy (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Sell (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Overcut (%)")}{row}'].number_format = '0.00%'
    ws[f'{L("Undercut (%)")}{row}'].number_format = '0.00%'
    ws[f'{L("Overcut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Undercut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Qty")}{row}'].number_format = '0'
    ws[f'{L("Theoretical Profit")}{row}'].number_format = '0.00'
    ws[f'{L("Amount Received")}{row}'].number_format = '0.00'
    ws[f'{L("ROI (%)")}{row}'].number_format = '0.00%'
    ws[f'{L("Demand-Supply Gap (%)")}{row}'].number_format = '0.00%'

    for int_col in ["Demand", "Supply", "Bought", "Sold", "Bids", "Offers"]:
        ws[f'{L(int_col)}{row}'].number_format = '#,##0'

    # Formulas
    ws[f'{L("Overcut (g)")}{row}'].value = f'={L("Buy (g.s)")}{row}*{L("Overcut (%)")}{row}'
    ws[f'{L("Undercut (g)")}{row}'].value = f'={L("Sell (g.s)")}{row}*{L("Undercut (%)")}{row}'
    ws[f'{L("Theoretical Profit")}{row}'].value = (
        f'=(({L("Undercut (g)")}{row}*0.85)-{L("Overcut (g)")}{row})*{L("Qty")}{row}'
    )
    ws[f'{L("Amount Received")}{row}'].value = f'={L("Undercut (g)")}{row}*{L("Qty")}{row}'
    ws[f'{L("ROI (%)")}{row}'].value = f'={L("Theoretical Profit")}{row}/({L("Overcut (g)")}{row}*{L("Qty")}{row})'
    ws[f'{L("Demand-Supply Gap (%)")}{row}'].value = f'=IF({L("Supply")}{row}=0,"",({L("Demand")}{row}-{L("Supply")}{row})/{L("Supply")}{row})'

# Auto-filter & width
ws.auto_filter.ref = ws.dimensions
for col_cells in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col_cells[0].column)
    for cell in col_cells:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(8, max_length + 2)

wb.save(output_file)
print(f"Final workbook saved to {output_file} with sheet '{safe_title}'.")
