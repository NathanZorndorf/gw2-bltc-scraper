import os
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import json

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
INPUT_FILE = "scraper-results.xlsx"
OUTPUT_FILE = "scraper-results-new.xlsx"

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
if os.path.exists(INPUT_FILE):
    existing_df = pd.read_excel(INPUT_FILE, sheet_name='scraper-results')
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

# Update columns to match scraper-formulas.csv
final_column_order = [
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers",
    "Overcut (%)", "Undercut (%)", "Overcut (g)", "Undercut (g)",
    "Max Flips / Day", "Bought/Bids", "Sold/Offers",
    "Buy-Through Rate (%)", "Sell-Through Rate (%)", "Flip-Through Rate (%)",
    "E(Profit | Qty = 1)", "E(ROI | Qty = 1)", "P(Buy = Qty)", "P(Sell = Qty)",
    "Optimal Qty", "Dynamic Sell-Through Rate (%)", "E(Sales | Q = Optimal Q)",
    "E(Profit | Q = Optimal Q)", "Optimal Investment (g)", "E(ROI | Q = Optimal Q)", "Time to Sell (Q Optimal)",
    "Target ROI", "Optimal Buy Price | Target ROI", "Optimal Qty | Target ROI",
    "Actual Qty Ordered", "Actual Buy Price", "Actual Sell Price",
    "Buy Order Placed", "Sell Order Placed", "Sold (manual)"
]

# Add new columns to DataFrame
for col in final_column_order:
    if col not in df.columns:
        df[col] = 0 if "Qty" in col or "Order" in col or "Sold" in col else ""

# Set default values for overcut/undercut columns using constants
df["Overcut (%)"] = OVERCUT_PCT_DEFAULT
df["Undercut (%)"] = UNDERCUT_PCT_DEFAULT
df["Target ROI"] = ROI_TARGET_DEFAULT

# Set default values for boolean columns
df["Buy Order Placed"] = False
df["Sell Order Placed"] = False
df["Sold (manual)"] = False

# Set default for Actual Qty Ordered
df["Actual Qty Ordered"] = ''
df["Actual Buy Price"] = ''
df["Actual Sell Price"] = ''

# --------------------
# COMBINE & SAVE
# --------------------
# Ensure all columns exist in both DataFrames
for col in final_column_order:
    if col not in existing_df.columns:
        existing_df[col] = ""
    if col not in df.columns:
        df[col] = ""

# Reorder columns
existing_df = existing_df[final_column_order]
df = df[final_column_order]

combined_df = pd.concat([existing_df, df], ignore_index=True)
combined_df = combined_df[final_column_order]  # Final re-order
combined_df.to_excel(OUTPUT_FILE, index=False)

# --------------------
# FORMAT EXCEL
# --------------------
wb = load_workbook(OUTPUT_FILE)
ws = wb.active
ws.title = "scraper-results"  # Use a fixed title for consistency
ws.freeze_panes = 'B2'


header_to_idx = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], start=1)}
def L(name): return get_column_letter(header_to_idx[name])
max_row = ws.max_row

for row in range(2, max_row + 1):
    # Number formats
    ws[f'{L("Buy (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Sell (g.s)")}{row}'].number_format = '0.00'
    ws[f'{L("Overcut (%)")}{row}'].number_format = '0%'
    ws[f'{L("Undercut (%)")}{row}'].number_format = '0%'
    ws[f'{L("Overcut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Undercut (g)")}{row}'].number_format = '0.00'
    ws[f'{L("Max Flips / Day")}{row}'].number_format = '0'
    ws[f'{L("Bought/Bids")}{row}'].number_format = '0.00'
    ws[f'{L("Sold/Offers")}{row}'].number_format = '0.00'
    ws[f'{L("Buy-Through Rate (%)")}{row}'].number_format = '0%'
    ws[f'{L("Sell-Through Rate (%)")}{row}'].number_format = '0%'
    ws[f'{L("Flip-Through Rate (%)")}{row}'].number_format = '0%'
    ws[f'{L("E(Profit | Qty = 1)")}{row}'].number_format = '0.00'
    ws[f'{L("E(ROI | Qty = 1)")}{row}'].number_format = '0%'
    ws[f'{L("P(Buy = Qty)")}{row}'].number_format = '0.00'
    ws[f'{L("P(Sell = Qty)")}{row}'].number_format = '0.00'
    ws[f'{L("Optimal Qty")}{row}'].number_format = '0'
    ws[f'{L("Dynamic Sell-Through Rate (%)")}{row}'].number_format = '0%'
    ws[f'{L("E(Sales | Q = Optimal Q)")}{row}'].number_format = '0.00'
    ws[f'{L("E(Profit | Q = Optimal Q)")}{row}'].number_format = '0.00'
    ws[f'{L("Optimal Investment (g)")}{row}'].number_format = '0.00'
    ws[f'{L("E(ROI | Q = Optimal Q)")}{row}'].number_format = '0%'
    ws[f'{L("Time to Sell (Q Optimal)")}{row}'].number_format = '0.00'
    ws[f'{L("Actual Qty Ordered")}{row}'].number_format = '0'
    ws[f'{L("Actual Buy Price")}{row}'].number_format = '0.00'

    # Formatting for new columns
    ws[f'{L("Target ROI")}{row}'].number_format = '0%'
    ws[f'{L("Optimal Qty | Target ROI")}{row}'].number_format = '0'
    ws[f'{L("Optimal Buy Price | Target ROI")}{row}'].number_format = '0.00'

    for int_col in ["Demand", "Supply", "Bought", "Sold", "Bids", "Offers"]:
        ws[f'{L(int_col)}{row}'].number_format = '#,##0'

    # Formulas
    ws[f'{L("Overcut (g)")}{row}'].value = f'={L("Buy (g.s)")}{row}*{L("Overcut (%)")}{row}'
    ws[f'{L("Undercut (g)")}{row}'].value = f'={L("Sell (g.s)")}{row}*{L("Undercut (%)")}{row}'
    ws[f'{L("Max Flips / Day")}{row}'].value = f'=MIN({L("Bought")}{row},{L("Sold")}{row})'
    ws[f'{L("Bought/Bids")}{row}'].value = f'=IFERROR({L("Bought")}{row}/{L("Bids")}{row},"")'
    ws[f'{L("Sold/Offers")}{row}'].value = f'=IFERROR({L("Sold")}{row}/{L("Offers")}{row},"")'
    ws[f'{L("Buy-Through Rate (%)")}{row}'].value = f'=IF({L("Bids")}{row}=0,IF({L("Bought")}{row}>0,1,0),MIN(1,{L("Bought")}{row}/{L("Bids")}{row}))'
    ws[f'{L("Sell-Through Rate (%)")}{row}'].value = f'=IF({L("Offers")}{row}=0,IF({L("Sold")}{row}>0,1,0),MIN(1,{L("Sold")}{row}/{L("Offers")}{row}))'
    ws[f'{L("Flip-Through Rate (%)")}{row}'].value = f'={L("Buy-Through Rate (%)")}{row}*{L("Sell-Through Rate (%)")}{row}'
    ws[f'{L("E(Profit | Qty = 1)")}{row}'].value = f'={L("Undercut (g)")}{row}*0.85*{L("Sell-Through Rate (%)")}{row}-{L("Overcut (g)")}{row}'
    ws[f'{L("E(ROI | Qty = 1)")}{row}'].value = f'=IFERROR({L("E(Profit | Qty = 1)")}{row}/{L("Overcut (g)")}{row},0)'
    ws[f'{L("P(Buy = Qty)")}{row}'].value = f'=BINOM.DIST({L("Actual Qty Ordered")}{row},{L("Actual Qty Ordered")}{row},{L("Buy-Through Rate (%)")}{row},FALSE)'
    ws[f'{L("P(Sell = Qty)")}{row}'].value = f'=BINOM.DIST({L("Actual Qty Ordered")}{row},{L("Actual Qty Ordered")}{row},{L("Sell-Through Rate (%)")}{row},FALSE)'
    ws[f'{L("Optimal Qty")}{row}'].value = f'=LET(q,ROUND(SQRT({L("Sold")}{row}*{L("Offers")}{row}*{L("Undercut (g)")}{row}*0.85/{L("Overcut (g)")}{row})-{L("Offers")}{row}),IF(q<0,0,MIN(q,{L("Max Flips / Day")}{row})))'
    ws[f'{L("Dynamic Sell-Through Rate (%)")}{row}'].value = f'=IFERROR(IF({L("Optimal Qty")}{row}>0,MIN(1,{L("Sold")}{row}/({L("Offers")}{row}+{L("Optimal Qty")}{row})),NA()),"")'
    ws[f'{L("E(Sales | Q = Optimal Q)")}{row}'].value = f'={L("Optimal Qty")}{row}*{L("Dynamic Sell-Through Rate (%)")}{row}'
    ws[f'{L("E(Profit | Q = Optimal Q)")}{row}'].value = f'={L("E(Sales | Q = Optimal Q)")}{row}*{L("Undercut (g)")}{row}*0.85-{L("Overcut (g)")}{row}*{L("Optimal Qty")}{row}'
    ws[f'{L("Optimal Investment (g)")}{row}'].value = f'={L("Optimal Qty")}{row}*{L("Overcut (g)")}{row}'
    ws[f'{L("E(ROI | Q = Optimal Q)")}{row}'].value = f'=IFERROR({L("E(Profit | Q = Optimal Q)")}{row}/{L("Optimal Investment (g)")}{row},0)'
    ws[f'{L("Time to Sell (Q Optimal)")}{row}'].value = f'=({L("Offers")}{row} + {L("Optimal Qty")}{row})/{L("Sold")}{row}'

    # New formulas for Target ROI
    ws[f'{L("Target ROI")}{row}'].value = f'={ROI_TARGET_DEFAULT}'
    ws[f'{L("Optimal Buy Price | Target ROI")}{row}'].value = f'=IFERROR(({L("Undercut (g)")}{row}*0.85)/(1+{L("Target ROI")}{row}), 0)'
    ws[f'{L("Optimal Qty | Target ROI")}{row}'].value = f'=LET(q,ROUND(SQRT({L("Sold")}{row}*{L("Offers")}{row}*{L("Undercut (g)")}{row}*0.85/{L("Optimal Buy Price | Target ROI")}{row}) - {L("Offers")}{row}), IF(q<0,0,MIN(q,{L("Max Flips / Day")}{row})))'

ws.auto_filter.ref = ws.dimensions
for col_cells in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col_cells[0].column)
    for cell in col_cells:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
    # ws.column_dimensions[col_letter].width = max(8, max_length + 2)
    ws.column_dimensions[col_letter].width = max_length

wb.save(OUTPUT_FILE)
print(f"Final workbook saved to {OUTPUT_FILE}.")