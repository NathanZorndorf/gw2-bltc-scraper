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
    "profit-pct-min": 10,
    "profit-pct-max": 100,
    "sold-day-min": 5,
    "bought-day-min": 5,
    "ipg": 200,
    "sort": "profit-pct",
    "page": 1
}

OVERCUT_PCT_DEFAULT = 1.10
UNDERCUT_PCT_DEFAULT = 0.90
ROI_TARGET_DEFAULT = 0.10
QTY_DEFAULT = 1
OUTPUT_FILE = "scraper-results.xlsx"

# Timestamp
scrape_time_dt = datetime.now()
scrape_time_str = scrape_time_dt.strftime("%Y-%m-%d %H:%M")

# --------------------
# HELPERS
# --------------------
def parse_gold_silver(td):
    gold = silver = 0
    for span in td.find_all("span"):
        classes = span.get("class", [])
        if "cur-t1c" in classes:
            gold = int(span.get_text(strip=True).replace(",", "") or 0)
        elif "cur-t1b" in classes:
            silver = int(span.get_text(strip=True) or 0)
    return round(gold + silver / 100, 2)

def parse_int(td):
    txt = td.get_text(strip=True).replace(",", "")
    try:
        return int(txt)
    except ValueError:
        return 0

# --------------------
# SCRAPE UNTIL EMPTY
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
        break

    for row in rows:
        cols = row.find_all("td")
        if len(cols) < 12:
            continue
        item_name = cols[1].get_text(strip=True)
        link_tag = cols[1].find("a", href=True)
        item_link = f"https://www.gw2bltc.com{link_tag['href']}" if link_tag else ""
        sell = parse_gold_silver(cols[2])
        buy = parse_gold_silver(cols[3])
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
# LOAD EXISTING FILE
# --------------------
if os.path.exists(OUTPUT_FILE):
    existing_df = pd.read_excel(OUTPUT_FILE)
    existing_df = existing_df[existing_df["Buy Order Placed"] == True]  # Keep only placed orders
else:
    existing_df = pd.DataFrame()

# --------------------
# CREATE NEW DATAFRAME
# --------------------
df = pd.DataFrame(all_rows, columns=[
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers"
])
df["Overcut (%)"] = OVERCUT_PCT_DEFAULT
df["Undercut (%)"] = UNDERCUT_PCT_DEFAULT
df["Overcut (g)"] = 0
df["Undercut (g)"] = 0
df["Qty"] = QTY_DEFAULT
df["Theoretical Profit - WF1"] = 0
df["Amount Received"] = 0
df["ROI (%)"] = 0
df["Demand-Supply Gap (%)"] = 0
df["ROI (Target %)"] = ROI_TARGET_DEFAULT
df["Bid / Item (g)"] = 0
df["Overcut (%) - WF2"] = 0
df["Offer / Item (g)"] = 0
df["Theoretical Profit - WF2"] = 0
df["Buy Order Placed"] = False
df["Sell Order Placed"] = False
df["Sold (manual)"] = False

final_column_order = [
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers",
    "Overcut (%)", "Undercut (%)", "Overcut (g)", "Undercut (g)",
    "Qty", "Theoretical Profit - WF1", "Amount Received", "ROI (%)",
    "Demand-Supply Gap (%)", "ROI (Target %)", "Bid / Item (g)", "Overcut (%) - WF2", "Offer / Item (g)",
    "Theoretical Profit - WF2",
    "Buy Order Placed", "Sell Order Placed", "Sold (manual)"
]
df = df[final_column_order]

# --------------------
# COMBINE & SAVE
# --------------------
combined_df = pd.concat([existing_df, df], ignore_index=True)
combined_df.to_excel(OUTPUT_FILE, index=False)

# --------------------
# FORMAT EXCEL
# --------------------
wb = load_workbook(OUTPUT_FILE)
ws = wb.active
safe_title = scrape_time_str.replace(":", "-")
ws.title = safe_title
ws.freeze_panes = 'B2'

header_to_idx = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], start=1)}
def L(name): return get_column_letter(header_to_idx[name])
max_row = ws.max_row

for row in range(2, max_row + 1):  # Format all rows, not just new ones
    ws[f'{L("Buy (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Sell (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Overcut (%)")}{row}'].number_format = '0%'
    ws[f'{L("Undercut (%)")}{row}'].number_format = '0%'
    ws[f'{L("Overcut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Undercut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Qty")}{row}'].number_format = '0'
    ws[f'{L("Theoretical Profit - WF1")}{row}'].number_format = '0.00'
    ws[f'{L("Amount Received")}{row}'].number_format = '0.00'
    ws[f'{L("ROI (%)")}{row}'].number_format = '0%'
    ws[f'{L("ROI (Target %)")}{row}'].number_format = '0%'
    ws[f'{L("Bid / Item (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Overcut (%) - WF2")}{row}'].number_format = '0%'
    ws[f'{L("Offer / Item (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Theoretical Profit - WF2")}{row}'].number_format = '0.00'
    ws[f'{L("Demand-Supply Gap (%)")}{row}'].number_format = '0%'

    for int_col in ["Demand", "Supply", "Bought", "Sold", "Bids", "Offers"]:
        ws[f'{L(int_col)}{row}'].number_format = '#,##0'

    ws[f'{L("Overcut (g)")}{row}'].value = f'={L("Buy (g.s)")}{row}*{L("Overcut (%)")}{row}'
    ws[f'{L("Undercut (g)")}{row}'].value = f'={L("Sell (g.s)")}{row}*{L("Undercut (%)")}{row}'
    ws[f'{L("Theoretical Profit - WF1")}{row}'].value = f'=(({L("Undercut (g)")}{row}*0.85)-{L("Overcut (g)")}{row})*{L("Qty")}{row}'
    ws[f'{L("Amount Received")}{row}'].value = f'={L("Undercut (g)")}{row}*{L("Qty")}{row}'
    ws[f'{L("ROI (%)")}{row}'].value = f'={L("Theoretical Profit - WF1")}{row}/({L("Overcut (g)")}{row}*{L("Qty")}{row})'
    ws[f'{L("Demand-Supply Gap (%)")}{row}'].value = f'=IF({L("Supply")}{row}=0,"",({L("Demand")}{row}-{L("Supply")}{row})/{L("Supply")}{row})'
    ws[f'{L("Bid / Item (g)")}{row}'].value = f'=({L("Undercut (%)")}{row}*0.85*{L("Sell (g.s)")}{row})/({L("ROI (Target %)")}{row}+1)'
    ws[f'{L("Overcut (%) - WF2")}{row}'].value = f'=({L("Bid / Item (g)")}{row})/({L("Buy (g.s)")}{row})'
    ws[f'{L("Offer / Item (g)")}{row}'].value = f'={L("Undercut (%)")}{row}*{L("Sell (g.s)")}{row}'
    ws[f'{L("Theoretical Profit - WF2")}{row}'].value = f'=(({L("Offer / Item (g)")}{row}*0.85)-{L("Bid / Item (g)")}{row})*{L("Qty")}{row}'

ws.auto_filter.ref = ws.dimensions
for col_cells in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col_cells[0].column)
    for cell in col_cells:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(8, max_length + 2)

wb.save(OUTPUT_FILE)
print(f"Final workbook saved to {OUTPUT_FILE} with sheet '{safe_title}'.")
print(f"Final workbook saved to {OUTPUT_FILE} with sheet '{safe_title}'.")
