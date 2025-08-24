# Guild Wars 2 Market History API Documentation

## Overview
This API provides access to historical market data for Guild Wars 2 items, including hourly, daily, and total summaries. Data is aggregated from in-game listings and transactions, with statistical fields calculated for each item and time period.

---

## Endpoints

### Hourly Data
- **GET /gw2/v2/history/hourly/json**
  - Returns hourly market history for items in JSON format.
- **GET /gw2/v2/history/hourly/csv**
  - Returns hourly market history for items in CSV format.

#### Example Request
```
GET /gw2/v2/history/hourly/json?itemID=12345&start=2025-08-01T00:00:00Z&end=2025-08-10T00:00:00Z
```

---

## Filtering & Query Parameters
- **itemID**: Filter results to a specific item.
- **start**: Start date/time for the data range (ISO string).
- **end**: End date/time for the data range (ISO string).
- **fields**: Specify which fields to include in the response.
- **project**: If true, only selected fields are returned.
- **filter**: If true, applies additional query filters.
- **beautify**: If true, formats JSON for human readability.
- **sorting**: Sort results by date (`new` for descending, `old` for ascending).

Filtering is handled by the `defaultPath` function, which parses query parameters and applies them to the database query.

---

## Field Calculations
All price and value fields are stored in copper (the smallest GW2 currency unit). To display as gold/silver/copper, use:

```js
gold = Math.floor(price / 10000)
silver = Math.floor((price % 10000) / 100)
copper = price % 100
```

### Field Definitions
| Field Name                | Description & Calculation Formula                                                                                   |
|-------------------------- |--------------------------------------------------------------------------------------------------------------------|
| count                     | Number of records aggregated for the period (e.g., per hour).                                                      |
| date                      | ISO date string for the record (hour).                                                                             |
| itemID                    | The item’s unique identifier.                                                                                      |
| type                      | Type/category of the record (if present).                                                                          |
| buy_delisted              | Number of buy listings removed during the period.                                                                  |
| buy_listed                | Number of buy listings added during the period.                                                                    |
| buy_sold                  | Number of buy listings resulting in a sale during the period.                                                      |
| buy_value                 | Total value of buy transactions: `buy_sold * buy_price` (summed over the period).                                  |
| buy_price                 | The lowest buy price at the time of aggregation.                                                                   |
| buy_price_avg             | Average of all buy prices in the period.                                                                           |
| buy_price_max             | Maximum buy price in the period.                                                                                   |
| buy_price_min             | Minimum buy price in the period.                                                                                   |
| buy_price_stdev           | Standard deviation of buy prices.                                                                                  |
| buy_quantity              | Total quantity of items in buy listings during the period.                                                         |
| buy_quantity_avg          | Average quantity per buy listing.                                                                                   |
| buy_quantity_max          | Maximum quantity in a buy listing.                                                                                 |
| buy_quantity_min          | Minimum quantity in a buy listing.                                                                                 |
| buy_quantity_stdev        | Standard deviation of buy quantities.                                                                              |
| sell_delisted             | Number of sell listings removed during the period.                                                                 |
| sell_listed               | Number of sell listings added during the period.                                                                   |
| sell_sold                 | Number of sell listings resulting in a sale during the period.                                                     |
| sell_value                | Total value of sell transactions: `sell_sold * sell_price` (summed over the period).                               |
| sell_price                | The lowest sell price at the time of aggregation.                                                                  |
| sell_price_avg            | Average of all sell prices in the period.                                                                          |
| sell_price_max            | Maximum sell price in the period.                                                                                  |
| sell_price_min            | Minimum sell price in the period.                                                                                  |
| sell_price_stdev          | Standard deviation of sell prices.                                                                                 |
| sell_quantity             | Total quantity of items in sell listings during the period.                                                        |
| sell_quantity_avg         | Average quantity per sell listing.                                                                                  |
| sell_quantity_max         | Maximum quantity in a sell listing.                                                                                |
| sell_quantity_min         | Minimum quantity in a sell listing.                                                                                |
| sell_quantity_stdev       | Standard deviation of sell quantities.                                                                             |

---

## How Data Is Aggregated
- Data is collected from in-game listings and transactions.
- For each hour, all relevant buy/sell listings are processed.
- Statistical fields (min, max, avg, stdev) are calculated from arrays of prices/quantities.
- Listed, delisted, and sold counts are determined by comparing changes in quantities between time intervals.

---

## Example Response (JSON)
```json
{
  "count": 1,
  "date": "2025-08-01T01:00:00Z",
  "itemID": 12345,
  "buy_price": 1234,
  "buy_price_avg": 1250,
  "buy_quantity": 10,
  "sell_price": 1300,
  "sell_price_avg": 1320,
  "sell_quantity": 8
  // ...other fields...
}
```
To display prices in gold/silver/copper, use the conversion above.

---

## Notes
- All endpoints support filtering and field selection via query parameters.
- Data is returned in copper; convert for display as needed.
- For more details on field calculations, see the code in `controller_listingsProcessing.ts` and related files.
