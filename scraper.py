import os
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --------------------
# CONFIG
# --------------------
GW2_API_BASE_URL = "https://api.guildwars2.com/v2"
DATAWARS2_API_BASE_URL = "https://api.datawars2.ie/gw2/v2/history/hourly/json"
INPUT_FILE = "scraper-results.xlsx"
OUTPUT_FILE = "scraper-results-new.xlsx"

# Timestamp
scrape_time_dt = datetime.now()
scrape_time_str = scrape_time_dt.strftime("%Y-%m-%d %H:%M")

# --------------------
# STEP 1: Fetch all item IDs from the Guild Wars 2 API
# --------------------
def get_all_item_ids():
    """Fetches all item IDs from the GW2 commerce listings endpoint."""
    try:
        response = requests.get(f"{GW2_API_BASE_URL}/commerce/listings")
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching item IDs: {e}")
        return []

print("Fetching all item IDs...")
all_item_ids = get_all_item_ids()
print(f"Found {len(all_item_ids)} items.")


if not all_item_ids:
    print("No item IDs found. Exiting.")
    exit()

# --------------------
# STEP 2: Filter item IDs based on price and quantity
# --------------------
def filter_item_ids(item_ids):
    """Filters item IDs based on buy/sell price and quantity."""
    filtered_ids = []
    print("Fetching prices for filtering...")
    for i in range(0, len(item_ids), 200):
        batch_ids = item_ids[i:i+200]
        print(f"Processing batch {i//200 + 1}/{len(item_ids)//200 + 1}")
        try:
            response = requests.get(f"{GW2_API_BASE_URL}/commerce/prices?ids={','.join(map(str, batch_ids))}")
            response.raise_for_status()
            for item in response.json():
                buy_price = item['buys']['unit_price']
                buy_quantity = item['buys']['quantity']
                sell_price = item['sells']['unit_price']
                sell_quantity = item['sells']['quantity']

                if (buy_price > 500 and buy_quantity > 10) and \
                   (sell_price > 500 and sell_quantity > 10):
                    filtered_ids.append(item['id'])
        except requests.exceptions.RequestException as e:
            print(f"Error fetching prices for batch {i//200}: {e}")
    return filtered_ids

print("Filtering item IDs...")
all_item_ids = filter_item_ids(all_item_ids)
print(f"Found {len(all_item_ids)} items after filtering.")

# --------------------
# STEP 3: Fetch item names in batches from the Guild Wars 2 API
# --------------------
def get_item_details(item_ids):
    """Fetches item details in batches of 200."""
    item_details = {}
    for i in range(0, len(item_ids), 200):
        batch_ids = item_ids[i:i+200]
        try:
            response = requests.get(f"{GW2_API_BASE_URL}/items?ids={','.join(map(str, batch_ids))}")
            response.raise_for_status()
            for item in response.json():
                item_details[item['id']] = item['name']
        except requests.exceptions.RequestException as e:
            print(f"Error fetching item details for batch {i//200}: {e}")
    return item_details

print("Fetching item names...")
item_names = get_item_details(all_item_ids)
print(f"Found names for {len(item_names)} items.")

# --------------------
# STEP 3: Fetch historical data from the DataWars2 API
# --------------------
def get_historical_data(item_id):
    """Fetches the last 7 days of historical data for a single item."""
    end_date = datetime.now(timezone.utc)
    start_date = end_date - timedelta(days=7)
    params = {
        "itemID": item_id,
        "start": start_date.isoformat().replace('+00:00', 'Z'),
        "end": end_date.isoformat().replace('+00:00', 'Z'),
    }
    try:
        response = requests.get(DATAWARS2_API_BASE_URL, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching historical data for item {item_id}: {e}")
        return None

print("Fetching historical data for each item...")
all_item_data = []
for i, item_id in enumerate(all_item_ids):
    print(f"Processing item {i+1}/{len(all_item_ids)} (ID: {item_id})")
    historical_data = get_historical_data(item_id)
    if historical_data:
        all_item_data.append({
            "item_id": item_id,
            "name": item_names.get(item_id, "Unknown"),
            "historical_data": historical_data
        })

print(f"Successfully fetched data for {len(all_item_data)} items.")

# --------------------
# STEP 4: Perform data processing and calculations
# --------------------
def get_current_prices(item_ids):
    """Fetches current buy and sell prices for a list of item IDs."""
    prices = {}
    for i in range(0, len(item_ids), 200):
        batch_ids = item_ids[i:i+200]
        try:
            response = requests.get(f"{GW2_API_BASE_URL}/commerce/prices?ids={','.join(map(str, batch_ids))}")
            response.raise_for_status()
            for item in response.json():
                prices[item['id']] = {
                    "buy": item['buys']['unit_price'],
                    "sell": item['sells']['unit_price']
                }
        except requests.exceptions.RequestException as e:
            print(f"Error fetching prices for batch {i//200}: {e}")
    return prices

print("Fetching current prices...")
current_prices = get_current_prices(all_item_ids)

processed_data = []
for item in all_item_data:
    item_id = item['item_id']
    item_name = item['name']
    historical_data = item['historical_data']

    # Construct gw2bltc.com URL
    # Replace spaces with hyphens and remove special characters
    url_name = ''.join(e for e in item_name if e.isalnum() or e.isspace()).replace(' ', '-')
    item_link = f"https://www.gw2bltc.com/en/item/{item_id}-{url_name}"

    # Get current prices
    price_info = current_prices.get(item_id, {})
    buy_price = price_info.get("buy", 0)
    sell_price = price_info.get("sell", 0)

    # Process historical data
    buy_price_max_list = [h['buy_price_max'] for h in historical_data if 'buy_price_max' in h]
    sell_price_min_list = [h['sell_price_min'] for h in historical_data if 'sell_price_min' in h]

    demand = sum(h.get('buy_quantity', 0) for h in historical_data)
    supply = sum(h.get('sell_quantity', 0) for h in historical_data)
    bought = sum(h.get('buy_sold', 0) for h in historical_data)
    sold = sum(h.get('sell_sold', 0) for h in historical_data)
    bids = sum(h.get('buy_listed', 0) for h in historical_data)
    offers = sum(h.get('sell_listed', 0) for h in historical_data)

    # Calculate median and std dev
    median_buy_price = np.median(buy_price_max_list) if buy_price_max_list else 0
    median_sell_price = np.median(sell_price_min_list) if sell_price_min_list else 0
    std_dev_buy_price = np.std(buy_price_max_list) if buy_price_max_list else 0
    std_dev_sell_price = np.std(sell_price_min_list) if sell_price_min_list else 0

    processed_data.append({
        "Item Name": item_name,
        "Item Link": item_link,
        "Date of Scrape": scrape_time_str,
        "Buy (g.s)": buy_price / 10000,
        "Sell (g.s)": sell_price / 10000,
        "Demand": demand,
        "Supply": supply,
        "Bought": bought,
        "Sold": sold,
        "Bids": bids,
        "Offers": offers,
        "Median Buy Price (g.s)": median_buy_price / 10000,
        "Median Sell Price (g.s)": median_sell_price / 10000,
        "Std Dev Buy Price (g.s)": std_dev_buy_price / 10000,
        "Std Dev Sell Price (g.s)": std_dev_sell_price / 10000,
    })

print(f"Successfully processed data for {len(processed_data)} items.")

# --------------------
# STEP 5: Integrate the new data into the existing spreadsheet structure
# --------------------
if not processed_data:
    print("No data processed. Exiting.")
    exit()

# Load existing file
if os.path.exists(INPUT_FILE):
    existing_df = pd.read_excel(INPUT_FILE, sheet_name='scraper-results')
    existing_df = existing_df[existing_df["Buy Order Placed"] == True]
else:
    existing_df = pd.DataFrame()

# Create new DataFrame
df = pd.DataFrame(processed_data)

# Define final column order, including new statistical columns
final_column_order = [
    "Item Name", "Item Link", "Date of Scrape", "Buy (g.s)", "Sell (g.s)",
    "Demand", "Supply", "Bought", "Sold", "Bids", "Offers",
    "Median Buy Price (g.s)", "Median Sell Price (g.s)",
    "Std Dev Buy Price (g.s)", "Std Dev Sell Price (g.s)",
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

# Add missing columns to the new DataFrame
for col in final_column_order:
    if col not in df.columns:
        df[col] = ""

# Set default values
df["Overcut (%)"] = 1.10
df["Undercut (%)"] = 0.90
df["Target ROI"] = 0.10
df["Buy Order Placed"] = False
df["Sell Order Placed"] = False
df["Sold (manual)"] = False
df["Actual Qty Ordered"] = ''
df["Actual Buy Price"] = ''
df["Actual Sell Price"] = ''

# Combine with existing data
for col in final_column_order:
    if col not in existing_df.columns:
        existing_df[col] = ""

combined_df = pd.concat([existing_df, df], ignore_index=True)
combined_df = combined_df[final_column_order]
combined_df.to_excel(OUTPUT_FILE, index=False)

# Format Excel
wb = load_workbook(OUTPUT_FILE)
ws = wb.active
ws.title = "scraper-results"
ws.freeze_panes = 'B2'

header_to_idx = {str(cell.value).strip(): idx for idx, cell in enumerate(ws[1], start=1)}
def L(name): return get_column_letter(header_to_idx[name])
max_row = ws.max_row

for row in range(2, max_row + 1):
    # Number formats
    for col_name in ["Buy (g.s)", "Sell (g.s)", "Median Buy Price (g.s)", "Median Sell Price (g.s)", "Std Dev Buy Price (g.s)", "Std Dev Sell Price (g.s)", "Overcut (g)", "Undercut (g)", "E(Profit | Qty = 1)", "Optimal Investment (g)", "E(Profit | Q = Optimal Q)", "Optimal Buy Price | Target ROI", "Actual Buy Price"]:
        if col_name in header_to_idx:
            ws[f'{L(col_name)}{row}'].number_format = '0.00'

    for col_name in ["Overcut (%)", "Undercut (%)", "Buy-Through Rate (%)", "Sell-Through Rate (%)", "Flip-Through Rate (%)", "E(ROI | Qty = 1)", "Dynamic Sell-Through Rate (%)", "E(ROI | Q = Optimal Q)", "Target ROI"]:
        if col_name in header_to_idx:
            ws[f'{L(col_name)}{row}'].number_format = '0%'

    for col_name in ["Demand", "Supply", "Bought", "Sold", "Bids", "Offers", "Max Flips / Day", "Optimal Qty", "Optimal Qty | Target ROI", "Actual Qty Ordered"]:
        if col_name in header_to_idx:
            ws[f'{L(col_name)}{row}'].number_format = '#,##0'

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
    ws[f'{L("Target ROI")}{row}'].value = f'=0.10'
    ws[f'{L("Optimal Buy Price | Target ROI")}{row}'].value = f'=IFERROR(({L("Undercut (g)")}{row}*0.85)/(1+{L("Target ROI")}{row}), 0)'
    ws[f'{L("Optimal Qty | Target ROI")}{row}'].value = f'=LET(q,ROUND(SQRT({L("Sold")}{row}*{L("Offers")}{row}*{L("Undercut (g)")}{row}*0.85/{L("Optimal Buy Price | Target ROI")}{row}) - {L("Offers")}{row}), IF(q<0,0,MIN(q,{L("Max Flips / Day")}{row})))'


ws.auto_filter.ref = ws.dimensions
for col_cells in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col_cells[0].column)
    for cell in col_cells:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(max_length, 15)

wb.save(OUTPUT_FILE)
print(f"Final workbook saved to {OUTPUT_FILE}.")