import os
import requests
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import plotly.io as pio
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
    return {iid: data for iid, data in agg.items() if data["bought_qty"] > 0 and data["sold_qty"] > 0}

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

    # --- Analytics Visualizations with Plotly ---
    status_callback("Generating interactive analytics visualizations...")

    # Set Plotly template
    pio.templates.default = "plotly_dark"

    # Prepare data
    top_profit_items = df[df["ROI (g.s)"] > 0].sort_values("ROI (g.s)", ascending=False).head(10)
    top_roi_items = df[df["ROI (%)"] > 0].sort_values("ROI (%)", ascending=False).head(10)

    # Chart 1: Pie Chart for Top Profitable Items
    fig1 = go.Figure(data=[go.Pie(
        labels=top_profit_items["Item Name"],
        values=top_profit_items["ROI (g.s)"],
        hole=.3,
        hovertemplate="<b>%{label}</b><br>Profit: %{value:.2f}g<br>%{percent}<extra></extra>"
    )])
    fig1.update_layout(title_text="Top 10 Profitable Items (by Gold)", height=450)

    # Chart 2: Bar Chart for Top Items by ROI (Gold)
    fig2 = go.Figure(go.Bar(
        y=top_profit_items["Item Name"],
        x=top_profit_items["ROI (g.s)"],
        orientation='h',
        hovertemplate="<b>%{y}</b><br>Profit: %{x:.2f}g<extra></extra>"
    ))
    fig2.update_layout(
        title_text="Top 10 Items by Profit (Gold)",
        yaxis=dict(autorange="reversed"),
        xaxis_title="Profit (Gold)",
        height=450
    )

    # Chart 3: Bar Chart for Top Items by ROI (%)
    fig3 = go.Figure(go.Bar(
        y=top_roi_items["Item Name"],
        x=top_roi_items["ROI (%)"],
        orientation='h',
        hovertemplate="<b>%{y}</b><br>ROI: %{x:.2%}<extra></extra>"
    ))
    fig3.update_layout(
        title_text="Top 10 Items by ROI (%)",
        yaxis=dict(autorange="reversed"),
        xaxis_title="ROI (%)",
        xaxis=dict(tickformat=".0%"),
        height=450
    )

    # Combine charts into a single HTML file with a grid layout
    report_html_path = os.path.join(output_dir, "interactive_report.html")
    with open(report_html_path, 'w') as f:
        f.write("""
        <html>
        <head>
            <title>Interactive Report</title>
            <style>
                body { background-color: #1a1a1a; color: #f0f0f0; font-family: sans-serif; }
                h1 { text-align: center; }
                .grid-container {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 20px;
                    padding: 20px;
                }
                .grid-item {
                    background-color: #2a2a2a;
                    border-radius: 8px;
                    padding: 15px;
                }
                .grid-item-span-2 {
                    grid-column: span 2;
                }
            </style>
        </head>
        <body>
            <h1>Transaction Analysis Report</h1>
            <div class="grid-container">
        """)

        # Embed charts into grid items
        f.write(f'<div class="grid-item grid-item-span-2">{fig1.to_html(full_html=False, include_plotlyjs="cdn")}</div>')
        f.write(f'<div class="grid-item">{fig2.to_html(full_html=False, include_plotlyjs=False)}</div>')
        f.write(f'<div class="grid-item">{fig3.to_html(full_html=False, include_plotlyjs=False)}</div>')

        f.write("""
            </div>
        </body>
        </html>
        """)

    status_callback(f"Interactive report saved to {report_html_path}")

def run_transaction_scraper(api_key: str, output_dir: str, status_callback=None, days: int = 30):
    if status_callback is None:
        status_callback = print

    if not api_key:
        status_callback("Error: API Key is missing.")
        return

    os.makedirs(output_dir, exist_ok=True)

    buys, sells = fetch_transactions(api_key, status_callback)
    buys = filter_last_n_days(buys, status_callback, n=days)
    sells = filter_last_n_days(sells, status_callback, n=days)

    if not buys:
        status_callback(f"No buy transactions found in the last {days} days.")
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
