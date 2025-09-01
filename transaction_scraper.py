import os
import requests
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
from dotenv import load_dotenv

def parse_coins_to_gold_silver(coins):
    gold = coins // 10000
    silver = (coins % 10000) // 100
    return round(gold + silver / 100, 2)

def fetch_all_transactions(endpoint, api_key, status_callback):
    headers = {"Authorization": f"Bearer {api_key}"}
    all_tx = []
    page = 0
    page_size = 200
    while True:
        url = f"{endpoint}?page={page}&page_size={page_size}"
        try:
            r = requests.get(url, headers=headers, timeout=20)
            r.raise_for_status()
            batch = r.json()
            if not batch:
                break
            all_tx.extend(batch)
            status_callback(f"Fetched page {page + 1} of transactions from {endpoint.split('/')[-1]}...")
            if len(batch) < page_size:
                break
            page += 1
        except requests.exceptions.RequestException as e:
            status_callback(f"Error fetching page {page} from {endpoint}: {e}")
            break
    return all_tx

def fetch_transactions(api_key, status_callback):
    base = "https://api.guildwars2.com/v2/commerce/transactions/history"
    status_callback("Fetching buy transactions...")
    buys = fetch_all_transactions(f"{base}/buys", api_key, status_callback)
    status_callback("Fetching sell transactions...")
    sells = fetch_all_transactions(f"{base}/sells", api_key, status_callback)
    return buys, sells

def get_item_names(item_ids, status_callback):
    names = {}
    ids = list(set(item_ids))
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        try:
            r = requests.get("https://api.guildwars2.com/v2/items", params={"ids": ",".join(map(str, batch))}, timeout=20)
            r.raise_for_status()
            for item in r.json():
                if isinstance(item, dict) and "id" in item and "name" in item:
                    names[item["id"]] = item["name"]
        except Exception as e:
            status_callback(f"Error fetching item names for batch {batch}: {e}")
    for iid in ids:
        if iid not in names:
            names[iid] = f"Item {iid}"
    return names

def filter_last_n_days(transactions, status_callback, date_field="purchased", n=30):
    cutoff = datetime.now() - timedelta(days=n)
    filtered = []
    skipped = 0
    for tx in transactions:
        date_str = tx.get(date_field) or tx.get("created")
        if not date_str:
            skipped += 1
            continue
        try:
            if date_str.endswith("Z"):
                tx_date = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S%z")
            else:
                if "+" in date_str:
                    main, tz = date_str.split("+")
                    tz = tz.replace(":", "")
                    date_str = main + "+" + tz
                tx_date = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S%z")
            if tx_date >= cutoff.replace(tzinfo=tx_date.tzinfo):
                filtered.append(tx)
        except Exception:
            skipped += 1
            continue
    status_callback(f"Filtered transactions: {len(filtered)} kept, {len(transactions) - len(filtered)} discarded.")
    return filtered

def aggregate_transactions(buys, sells):
    agg = {}
    for tx in buys:
        iid = tx["item_id"]
        price = parse_coins_to_gold_silver(tx["price"])
        qty = tx["quantity"]
        spent = price * qty
        if iid not in agg:
            agg[iid] = {"bought_qty": 0, "spent": 0, "sold_qty": 0, "received": 0}
        agg[iid]["bought_qty"] += qty
        agg[iid]["spent"] += spent
    for tx in sells:
        iid = tx["item_id"]
        price = parse_coins_to_gold_silver(tx["price"])
        qty = tx["quantity"]
        received = price * qty
        if iid not in agg:
            continue
        agg[iid]["sold_qty"] += qty
        agg[iid]["received"] += received
    return {iid: data for iid, data in agg.items() if data["bought_qty"] > 0}

def save_profit_report(agg, item_names, output_dir, status_callback):
    output_file = os.path.join(output_dir, "profit-report.xlsx")
    rows = []
    for iid, data in agg.items():
        name = item_names.get(iid, f"Item {iid}")
        spent = data["spent"]
        received = data["received"]
        roi = received - spent
        roi_pct = roi / spent if spent else ""
        rows.append({
            "Item Name": name, "Bought Qty": data["bought_qty"], "Sold Qty": data["sold_qty"],
            "Total Spent (g.s)": spent, "Total Received (g.s)": received,
            "ROI (g.s)": roi, "ROI (%)": roi_pct
        })
    df = pd.DataFrame(rows)
    if df.empty or "ROI (g.s)" not in df.columns:
        status_callback("No transactions to report for the selected period/items.")
        return

    df["ROI (%)"] = pd.to_numeric(df["ROI (%)"], errors="coerce").fillna(0)
    df = df.sort_values("ROI (g.s)", ascending=False)
    df.to_excel(output_file, index=False)

    status_callback(f"Profit report saved to {output_file}")

    # --- Analytics Visualizations ---
    status_callback("Generating analytics visualizations...")
    img_dir = os.path.join(output_dir, "img")
    os.makedirs(img_dir, exist_ok=True)

    # ... (charting code remains the same, using img_dir)
    plt.ioff() # Turn off interactive mode

    top_profit_items = df[df["ROI (g.s)"] > 0].sort_values("ROI (g.s)", ascending=False).head(10)
    plt.figure(figsize=(8,8))
    plt.pie(top_profit_items["ROI (g.s)"], labels=top_profit_items["Item Name"], autopct='%1.1f%%', startangle=140)
    plt.title("Top 10 Profitable Items (ROI in Gold)")
    plt.tight_layout()
    plt.savefig(os.path.join(img_dir, "profit_pie_chart.png"))
    plt.close()

    status_callback("Analytics charts saved.")

def run_transaction_scraper(api_key: str, output_dir: str, status_callback=None):
    if status_callback is None:
        status_callback = print

    if not api_key:
        status_callback("Error: API Key is missing.")
        return

    os.makedirs(output_dir, exist_ok=True)

    buys, sells = fetch_transactions(api_key, status_callback)
    buys = filter_last_n_days(buys, status_callback, n=30)
    sells = filter_last_n_days(sells, status_callback, n=30)

    if not buys:
        status_callback("No buy transactions found in the last 30 days.")
        return

    all_ids = [tx["item_id"] for tx in buys + sells]
    status_callback("Fetching item names...")
    item_names = get_item_names(all_ids, status_callback)

    status_callback("Aggregating transactions...")
    agg = aggregate_transactions(buys, sells)

    save_profit_report(agg, item_names, output_dir, status_callback)
    status_callback("Transaction report complete.")


if __name__ == "__main__":
    load_dotenv()
    api_key = os.environ.get("GW2_API_KEY")
    output_dir = "."

    run_transaction_scraper(api_key=api_key, output_dir=output_dir)
