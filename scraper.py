import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

# --------------------
# CONFIG
# --------------------
BASE_URL = "https://www.gw2bltc.com/en/tp/search"
PARAMS = {
    "profit-min": 500,
    "profit-pct-min": 0,
    "profit-pct-max": 100,
    "sold-day-min": 10,
    "bought-day-min": 10,
    "ipg": 200,
    "sort": "profit-pct",
    "page": 1
}

OVERCUT_PCT_DEFAULT = 1.10
UNDERCUT_PCT_DEFAULT = 0.90
QTY_DEFAULT = 1

# Timestamp
scrape_time_dt = datetime.now()
scrape_time_str = scrape_time_dt.strftime("%Y-%m-%d %H:%M")

output_file = "gw2_trading_post_sorted.xlsx"

# --------------------
# HELPERS
# --------------------
def parse_gold_silver(td):
    gold = silver = 0
    for span in td.find_all("span"):
        classes = span.get("class", [])
        if "cur-t1c" in classes:
            try:
                gold = int(span.get_text(strip=True))
            except ValueError:
                gold = 0
        elif "cur-t1b" in classes:
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

def get_total_pages(soup):
    pagination_buttons = soup.select('.btn-group.btn-group-justified a.btn')
    if not pagination_buttons:
        return 1
    max_page = 1
    for button in pagination_buttons:
        href = button.get('href', '')
        if 'page=' in href:
            page_match = re.search(r'page=(\d+)', href)
            if page_match:
                max_page = max(max_page, int(page_match.group(1)))
        button_text = button.get_text(strip=True)
        if button_text.isdigit():
            max_page = max(max_page, int(button_text))
    return max_page

# --------------------
# PRE-SCRAPE CHECK
# --------------------
print("Checking total number of pages...")
try:
    first_page_r = requests.get(BASE_URL, params=PARAMS, timeout=20)
    first_page_r.raise_for_status()
    first_page_soup = BeautifulSoup(first_page_r.text, "html.parser")
    total_pages = get_total_pages(first_page_soup)
    print(f"Found {total_pages} pages.")
except requests.exceptions.RequestException as e:
    print(f"Error fetching page data: {e}")
    exit()

# --------------------
# SCRAPE
# --------------------
all_rows = []

while PARAMS["page"] <= total_pages:
    print(f"Fetching page {PARAMS['page']}/{total_pages}...")
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
# DATAFRAME
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
df["Theoretical Profit"] = 0
df["Amount Received"] = 0
df["ROI (%)"] = 0
df["Demand-Supply Gap (%)"] = 0
df["Buy Order Placed"] = False
df["Sell Order Placed"] = False
df["Sold (manual)"] = False

final_column_order = [
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers",
    "Overcut (%)", "Undercut (%)", "Overcut (g)", "Undercut (g)",
    "Qty", "Theoretical Profit", "Amount Received", "ROI (%)",
    "Demand-Supply Gap (%)", "Buy Order Placed", "Sell Order Placed", "Sold (manual)"
]
df = df[final_column_order]

df.to_excel(output_file, index=False)

# --------------------
# FORMAT EXCEL
# --------------------
wb = load_workbook(output_file)
ws = wb.active
ws.title = scrape_time_str
ws.freeze_panes = 'B2'

max_row = ws.max_row
header_to_idx = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], start=1)}
def L(name): return get_column_letter(header_to_idx[name])

for row in range(2, max_row + 1):
    # Formats
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
print(f"Final workbook saved to {output_file} with sheet '{scrape_time_str}'.")
