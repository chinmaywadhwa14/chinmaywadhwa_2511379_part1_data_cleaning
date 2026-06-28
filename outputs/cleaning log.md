# Cleaning Log — Retail Order Dataset

## 1. List of Issues Found

| # | Issue | Field(s) | Records Affected |
|---|---|---|---|
| 1 | Inconsistent case, extra/leading/trailing spaces, double spaces | customer_name, segment, region, state, city, category, sub_category, ship_mode, payment_status, order_status | ~60 records with visible formatting issues (every record was reformatted defensively) |
| 2 | Four different date formats mixed in the same column | order_date, ship_date | All 912 records (formats: `DD Mon YYYY`, `MM/DD/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD`) |
| 3 | Ship date earlier than order date | order_date, ship_date | 21 records |
| 4 | Exact duplicate rows (all 21 fields identical) | entire row | 40 rows (20 extra copies of 20 order_ids) |
| 5 | Duplicate order_id with conflicting field values | order_id + others | 12 order_ids / 24 records |
| 6 | Missing values | region | 25 records |
| 7 | Missing values | ship_mode | 21 records |
| 8 | Missing values | discount | 18 records |
| 9 | Discount stored as percentage text (e.g. "70%") instead of a decimal | discount | 8 records |
| 10 | Negative discount values | discount | 15 records |
| 11 | Discount values above the allowed range | discount | 15 records |
| 12 | Sales value does not match quantity × unit_price × (1-discount) | sales | 64 records |
| 13 | Profit value does not match sales − cost | profit | 64 records |
| 14 | Order marked Completed but payment_status is Failed (logical inconsistency) | order_status, payment_status | 2 records |

No missing or unrecognized-text values were found in order_id, order_date, ship_date, customer_id, customer_name, quantity, unit_price, sales, cost, profit, payment_status, or order_status.

## 2. Cleaning Actions Performed

- **Text fields** — standardized with `=PROPER(TRIM(...))` to fix case inconsistencies and collapse/remove extra spaces, applied via live formulas in `cleaned_orders.xlsx` (sheet `cleaned_orders`), referencing an internal `raw_data` sheet that holds the de-duplicated raw values.
- **Dates** — a single nested `IF`/`MID`/`DATE`/`MATCH` formula detects which of the 4 raw formats each cell uses and converts it to a true Excel date, displayed consistently as `dd-mmm-yyyy`. Formula is in `order_date` and `ship_date` columns.
- **Discount** — standardized with a formula that converts percentage-text values to decimals, defaults blanks to 0, and passes numeric values through unchanged (`cleaned_discount` column).
- **Duplicates** — exact duplicate rows were removed before any other processing (20 rows). Conflicting duplicates were kept and flagged (see Section 3).
- **Calculated columns** — `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year` were all added as live formulas (not hardcoded), so the workbook recalculates automatically if source data changes.

## 3. Business Rules Applied

| Rule Area | Action Taken |
|---|---|
| Missing region | Filled as "Unknown"; flagged in `region_filled_flag` |
| Missing ship_mode | Filled as "Unknown"; flagged in `ship_mode_filled_flag` |
| Missing discount | Treated as 0 (all other sales fields were valid for these records) |
| Negative discount | Flagged as invalid in `discount_issue_flag`; original value preserved in `discount_raw` |
| Discount above allowed range | Flagged as invalid in `discount_issue_flag`; original value preserved |
| Cancelled orders | Excluded from the completed sales summary (`excluded_from_sales_summary` = Excluded) |
| Failed payments | Excluded from the completed sales summary, even if order_status = Completed |
| Refunded orders | Kept in the dataset and summarized separately (`refund_flag`, and Pivot 5 in `pivot_summary.xlsx`) |
| Ship date before order date | Flagged as an invalid shipping record (`ship_before_order_flag`); record is kept, not deleted |
| Exact duplicate rows | First occurrence kept, extra copies removed (not silently — logged in `data_quality_report.xlsx`) |
| Conflicting duplicate order_id | Both/all records kept and flagged for manual review — never silently deleted |

## 4. Assumptions Made

1. **Allowed discount range = 0%–30%.** The dataset's standard discount tiers are 0%, 5%, 10%, 15%, 20%, and 25%. Any discount outside 0–30% (including the observed 55%, 65%, 70%, 85% values and all negative values) is treated as invalid.
2. **"Completed sales summary" = order_status is "Completed" AND payment_status is not "Failed."** This satisfies both stated rules ("Cancelled should not contribute" and "Failed payments should not contribute") simultaneously and also screens out the 2 records that were marked Completed despite a Failed payment, which is treated as a data anomaly rather than a genuine completed sale. Returned orders are not included in the completed sales pivots, since "completed sales" should reflect orders that were actually fulfilled and paid for.
3. **MM/DD/YYYY vs DD-MM-YYYY format detection** was determined by checking which position in each format consistently exceeds 12 across the dataset (confirmed: slash-format day value never exceeds 12 in the first position and the second position goes up to 31, so slash dates are MM/DD/YYYY; dash-format is the reverse, so dash dates are DD-MM-YYYY).
4. **Percentage-text discounts** (e.g., "70%") are assumed to represent the literal percentage written (i.e., divided by 100), not a typo for a smaller decimal.
5. A "mismatch" between a raw and recalculated sales/profit value is only flagged if the difference exceeds ₹1.00, to avoid flagging trivial floating-point rounding differences.

## 5. Records Removed

- **20 rows** removed — exact duplicate copies (all 21 original fields identical to another row). The first occurrence of each was always kept. Full list of removed order_ids and their original raw row numbers is documented in `data_quality_report.xlsx` → "2. Duplicate Summary".
- **No other records were removed.** Every other row from the raw file, including every conflicting duplicate, every missing-value record, and every invalid-discount record, is present in `cleaned_orders.xlsx`.

## 6. Records Flagged

Out of 912 cleaned records:
- **505 Clean** — no issues of any kind.
- **295 Warning** — a minor, corrected issue (Unknown-filled region/ship_mode, discount defaulted to 0, a refunded order, or order_status Cancelled/Returned).
- **112 Invalid** — a material data integrity issue requiring review (invalid discount, ship date before order date, a sales/profit calculation mismatch, or a conflicting duplicate order_id). See `data_quality_report.xlsx` for the full breakdown and record-level detail in every section.

## 7. Limitations of the Cleaning Process

- Conflicting duplicate order_ids (12 order_ids / 24 records) were **flagged, not resolved** — a human reviewer with access to the source systems is needed to determine which version of each conflicting record is correct.
- Sales/profit mismatches (64 + 64 records) were **flagged, not corrected** — the original `sales_raw`/`profit_raw` values are preserved alongside the recalculated values so a reviewer can decide which is authoritative; the recalculated values are used in all pivot/summary reporting.
- The discount allowed-range assumption (0–30%) is a reasonable inference from the data, not a value confirmed by the business; it should be validated with stakeholders before being treated as a hard rule.
- Date parsing assumes only the 4 formats observed in this export; a future export containing a fifth format would need the parsing formula extended.
- "Unknown" region/ship_mode placeholders mean some pivot summaries include a small "Unknown" bucket that cannot be reassigned without the original source data.
