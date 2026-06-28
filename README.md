# Part 1 — Business Data Cleaning, Validation & Excel Reporting

## 1. Problem Summary

As a business analyst for a retail company, the task was to take a messy, order-level sales export from multiple internal systems and turn it into a clean, validated, analysis-ready dataset, along with a set of summary reports a manager could use for a business review. The raw export contained inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discounts, sales/profit calculation mismatches, and order-status inconsistencies that all needed to be identified, documented, and resolved according to a defined set of business rules.

## 2. Dataset Description

- **Source file:** `data/raw_orders.xlsx` — 932 order-level records, 21 columns (order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status).
- **Cleaned file:** `data/cleaned_orders.xlsx` — 912 records (20 exact-duplicate rows removed), 40 columns: every original field cleaned/standardized plus calculated columns (`calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`) and data-quality flag columns.

## 3. Tools Used

- **Microsoft Excel / openpyxl-built workbooks** — all cleaning, validation, and reporting logic is implemented as native Excel formulas (TRIM, PROPER, SUBSTITUTE-style text cleaning, DATE/MID/MATCH-based date parsing, COUNTIFS/SUMIFS/AVERAGEIFS pivot-style aggregation) so every output recalculates automatically and contains zero hardcoded calculated values.
- **LibreOffice headless recalculation** — used to force a full formula recalculation and confirm zero formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?) across all three workbooks before delivery.

## 4. Cleaning Steps Performed

1. Removed 20 exact-duplicate rows (all 21 fields identical) — kept the first occurrence of each.
2. Standardized all text fields (customer_name, segment, region, state, city, category, sub_category, ship_mode, payment_status, order_status) using TRIM + PROPER logic to fix case and spacing inconsistencies.
3. Parsed and standardized order_date and ship_date from 4 mixed formats into one consistent date format, and flagged any ship date earlier than its order date.
4. Standardized the discount field — converted percentage-text values (e.g. "70%") to decimals, defaulted missing values to 0, and flagged negative or out-of-range values as invalid.
5. Recalculated sales and profit from quantity, unit_price, cleaned discount, and cost, and flagged every record where the raw value didn't match the recalculation by more than ₹1.
6. Added calculated columns: `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`.
7. Built an overall `data_quality_flag` (Clean / Warning / Invalid) per record summarizing every issue found.

See `outputs/cleaning_log.md` for full detail.

## 5. Business Rules Applied

| Rule Area | Action Taken |
|---|---|
| Missing region / ship_mode | Filled as "Unknown", flagged |
| Missing discount | Treated as 0 (only where other sales fields were valid) |
| Negative / out-of-range discount | Flagged as invalid |
| Cancelled orders / Failed payments | Excluded from the completed sales summary |
| Refunded orders | Kept and summarized separately |
| Ship date before order date | Flagged as an invalid shipping record |
| Duplicate handling | Exact duplicates removed; conflicting duplicates kept and flagged for review (never silently deleted) |

## 6. Summary of Data Quality Issues Found

Out of 912 cleaned records: **505 Clean**, **295 Warning**, **112 Invalid**. Full breakdown — missing values, duplicates, invalid discounts, date issues, order-status issues, and calculation mismatches — is in `outputs/data_quality_report.xlsx` (7 sections) and `outputs/cleaning_log.md`.

## 7. Summary of Final Pivot Reports

`outputs/pivot_summary.xlsx` contains 6 pivot-style summaries (all built with live SUMIFS/COUNTIFS/AVERAGEIFS formulas, two of which are sorted/filtered):

1. Sales and profit by region (sorted by total sales, descending)
2. Sales and profit by category and sub-category
3. Order count by ship mode (sorted descending, AutoFilter enabled)
4. Profit margin by customer segment
5. Refunded / cancelled / failed orders by region
6. Monthly sales trend (AutoFilter enabled)

## 8. Key Business Insights

- **South and West** generate the highest completed sales, but **West has the strongest average profit margin (~30%)** among the major regions — worth a closer look at why South's margin trails despite leading on revenue.
- **Standard Class** is the most-used ship mode (27% of orders) but also has the longest average shipping delay (~4 days), which is worth investigating from a customer-experience standpoint.
- **Home Office and Small Business** segments lead on total sales and profit margin; **Consumer** has the lowest average margin (~22%) of the four segments.
- **112 of 912 records (12%)** carry a material data-quality issue (invalid discount, date inconsistency, calculation mismatch, or conflicting duplicate) that should be resolved with the source systems before this data is used for financial reporting.
- **64 records** have a sales or profit figure that doesn't match a recalculation from quantity, unit price, discount, and cost — this is the single largest data-integrity risk in the dataset and warrants a root-cause investigation with IT/source-system owners.

## 9. Assumptions and Limitations

- Valid discount range assumed to be 0%–30%, inferred from the dataset's own standard tiers (0/5/10/15/20/25%).
- "Completed sales" defined as order_status = Completed AND payment_status ≠ Failed.
- Conflicting duplicate order_ids and sales/profit mismatches were flagged, not resolved — they require a human reviewer with access to the source systems.
- Full list of assumptions and limitations is in `outputs/cleaning_log.md` (Sections 4 and 7).

## 10. Screenshots Included

- `screenshots/raw_data_preview.png` — raw dataset before cleaning
- `screenshots/cleaned_data_preview.png` — cleaned dataset with calculated columns
- `screenshots/pivot_summary_1.png` — a major pivot summary
- `screenshots/pivot_summary_2.png` — another major pivot summary
