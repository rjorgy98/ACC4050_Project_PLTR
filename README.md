# ACC4050_Project_PLTR

## Financial statement data locations (Palantir_Financials.xlsx)
The core statement tables start after the header blocks in each sheet:

- **INCOME_STATEMENT**: headers in row 15 with years in columns C–E; line items run from B16:B42 with values in C16:E42.
- **BALANCE_SHEET**: headers in row 14 with years in columns C–D; line items run from B17:B55 with values in C17:D55.
- **STOCKHOLDERS_EQUITY**: headers in row 17 with line items running from B18:B62 and values in C18:J62.
- **CASH_FLOW**: headers in row 15 with years in columns C–E; line items run from B17:B68 with values in C17:E68.

## Ratio workflow (reusable for similar workbooks)
1. **Confirm the statement ranges** match the layout above (header row with years, label column, and data rows). Update `SHEET_CONFIGS` in `ratio_calculator.py` if the workbook differs.
2. **Install dependencies** if needed:
   ```bash
   pip install openpyxl
   ```
3. **Run the ratio script** to generate the `RATIOS` sheet:
   ```bash
   python ratio_calculator.py
   ```
4. **Review `RATIOS`** for the computed values. Inventory-based ratios will be blank if no inventory line exists.

The script calculates A/R Turnover, Inventory Turnover (if applicable), A/P Turnover, PPE Turnover, Asset Turnover, Return on Assets, Cash Conversion Cycle, Return on Equity, Return on Common Equity, Net Profit Margin, and Leverage.
