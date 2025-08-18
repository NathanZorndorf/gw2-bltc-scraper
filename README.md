# Scraper Formulas Explanation

This document explains the formulas used in `scraper.py` for the new columns related to Target ROI.

### `Optimal Buy Price | Target ROI`

This column calculates the maximum price you should pay for an item to achieve a specific `Target ROI`. The formula is derived from the standard Return on Investment (ROI) formula.

The basic ROI formula is:
```
ROI = (Revenue - Cost) / Cost
```

In the context of the Guild Wars 2 Trading Post, we have the following:
- **Revenue**: The selling price of the item, minus the 15% transaction fee. This is represented as `Sell_Price * 0.85`. In the Excel sheet, the `Undercut (g)` column is used as the `Sell_Price`.
- **Cost**: The price at which you buy the item. This is the value we want to calculate. Let's call it `Buy_Price`.

So the formula becomes:
```
Target_ROI = (Sell_Price * 0.85 - Buy_Price) / Buy_Price
```

To find the `Buy_Price` that will give us our `Target_ROI`, we can rearrange the formula:

1. Multiply both sides by `Buy_Price`:
   ```
   Target_ROI * Buy_Price = Sell_Price * 0.85 - Buy_Price
   ```

2. Add `Buy_Price` to both sides:
   ```
   Target_ROI * Buy_Price + Buy_Price = Sell_Price * 0.85
   ```

3. Factor out `Buy_Price`:
   ```
   Buy_Price * (Target_ROI + 1) = Sell_Price * 0.85
   ```

4. Divide by `(Target_ROI + 1)` to solve for `Buy_Price`:
   ```
   Buy_Price = (Sell_Price * 0.85) / (1 + Target_ROI)
   ```

This gives us the final Excel formula for `Optimal Buy Price | Target ROI`:
```excel
=IFERROR(({Undercut (g)}*0.85)/(1+{Target ROI}), 0)
```
The `IFERROR` is used to prevent errors if the input values are not valid.

### `Optimal Qty | Target ROI`

This column calculates the optimal quantity of an item to trade, given the `Optimal Buy Price | Target ROI` that we just calculated. The formula is based on the existing formula for `Optimal Qty`, but with the new buy price.

The original `Optimal Qty` formula is designed to maximize profit, and it is:
```excel
=LET(q,ROUND(SQRT({Sold}*{Offers}*{Undercut (g)}*0.85/{Overcut (g)})-{Offers}),IF(q<0,0,MIN(q,{Max Flips / Day})))
```
This formula is derived from an economic model where the profit is maximized. The `{Overcut (g)}` term represents the buy price.

To calculate the `Optimal Qty` for our `Target ROI`, we simply substitute the original buy price (`{Overcut (g)}`) with our new calculated `Optimal Buy Price | Target ROI`.

This gives the following formula for `Optimal Qty | Target ROI`:
```excel
=IF({Optimal Buy Price | Target ROI}=0, 0, LET(q,ROUND(SQRT({Sold}*{Offers}*{Undercut (g)}*0.85/{Optimal Buy Price | Target ROI}) - {Offers}), IF(q<0,0,MIN(q,{Max Flips / Day}))))
```
This formula calculates the optimal quantity to trade if you were to buy at the price that guarantees your `Target ROI`. An `IF` condition is added to handle cases where the `Optimal Buy Price | Target ROI` is zero, to avoid division by zero errors.
