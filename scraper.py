from bs4 import BeautifulSoup
from datetime import datetime
import csv

# Load the saved HTML file
with open("Trading Post _ GW2BLTC.html", "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f, "html.parser")

# Output file
output_csv = "gw2_trading_post.csv"

# Current date for scrape date column
scrape_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Find all table rows after the header
rows = soup.select("table.table-result tr")[1:]  # skip the header row

# Prepare CSV
headers = [
    "Item Name",
    "Item Link",
    "Date of Scrape",
    "Sell (g.s)",
    "Buy (g.s)",
    "Supply",
    "Demand",
    "Sold",
    "Offers",
    "Bought",
    "Bids"
]
data = []

for row in rows:
    cols = row.find_all("td")
    if len(cols) < 12:  # skip rows that don't match the data format
        continue

    # Item name
    item_name = cols[1].get_text(strip=True)
    # Item link (absolute URL)
    link_tag = cols[1].find("a", href=True)
    item_link = f"https://www.gw2bltc.com{link_tag['href']}" if link_tag else ""

    # Helper to parse gold/silver into decimal format, ignoring copper
    def parse_gold_silver(td):
        gold = silver = 0
        for span in td.find_all("span"):
            classes = span.get("class", [])
            if "cur-t1c" in classes:
                gold = int(span.get_text(strip=True))
            elif "cur-t1b" in classes:
                silver = int(span.get_text(strip=True))
        return round(gold + silver / 100, 2)

    sell = parse_gold_silver(cols[2])
    buy = parse_gold_silver(cols[3])

    # Remove commas and convert to integers for numeric columns
    def parse_int(td):
        return int(td.get_text(strip=True).replace(",", ""))

    supply = parse_int(cols[6])
    demand = parse_int(cols[7])
    sold = parse_int(cols[8])
    offers = parse_int(cols[9])
    bought = parse_int(cols[10])
    bids = parse_int(cols[11])

    data.append([
        item_name,
        item_link,
        scrape_date,
        sell,
        buy,
        supply,
        demand,
        sold,
        offers,
        bought,
        bids
    ])

# Write to CSV
with open(output_csv, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    writer.writerows(data)

print(f"CSV saved as {output_csv} with {len(data)} items.")
