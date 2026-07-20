# Retail Sales Data Cleaning & Analysis

A self-directed Excel practice project — raw data cleaning, PivotTable summarization, and interactive dashboard reporting.

## Overview

This workbook takes a deliberately messy 148-row retail sales export and turns it into a clean, analysis-ready dataset with a summary dashboard. It's a practice dataset (not real business data), built to demonstrate a full data-cleaning-to-reporting workflow in Excel: identifying data quality issues, resolving them with a documented and defensible method, and summarizing results with a native PivotTable, slicers, and charts.

**[Download the workbook](./Retail_Sales_Practice_Project.xlsx)**

## What's in the file

| Tab | Contents |
|---|---|
| **Raw Data** | The sales export exactly as received — inconsistent date formats, inconsistent text casing/spacing, missing values, and duplicate orders |
| **Clean Data** | The same 148 rows after cleaning, reduced to 134 verified records, with a `Date_Confidence` column flagging any assumption made during date cleanup |
| **Pivot Summary** | An interactive dashboard — total sales and order counts by Region, Category, and Month, each with a chart, plus slicers for filtering by Region, Category, Month, and Payment Method |

## Data cleaning process

1. Standardized three mixed date formats (ISO, DD/MM/YYYY-style, and "Mon DD, YYYY" text) into one consistent format
2. Removed stray leading/trailing whitespace from text fields
3. Standardized inconsistent capitalization and spacing across Region and Category (e.g. `lagos `, `LAGOS` → `Lagos`)
4. Removed 8 exact duplicate order rows, matched on Order ID
5. Dropped 6 rows with a missing Quantity or Unit Price, rather than estimating a figure that would misstate revenue
6. Rebuilt Total as a live formula (`Quantity × Unit Price`) instead of a static number
7. Summarized the cleaned data with a native Excel PivotTable, added slicers for interactive filtering, and built a chart for each breakdown

## Handling ambiguous dates

Some dates in the raw export used a slash format where neither number is unambiguously a day or a month (e.g. `07/01/2026` could mean 7 January or 1 July). Rather than guess silently, every date was checked: if one of the two numbers is greater than 12, the format is provable from the text itself. Where both numbers were 12 or under, the date was genuinely ambiguous — those cases were resolved using the Day/Month convention confirmed by the unambiguous dates elsewhere in the same export, and each one is labeled `Assumed` rather than `Confirmed` in the `Date_Confidence` column so the assumption stays visible instead of hidden.

**Result:** 118 of 134 dates are `Confirmed` directly from the source text; 16 are `Assumed` under the documented rule above and flagged for review.

## Row counts

- Raw rows received: 148
- Duplicate rows removed: 8
- Rows dropped for missing Quantity/Unit Price: 6
- Clean rows analyzed: 134

## Key findings

- Lagos was the top-performing region (₦960,300 across 40 orders); Abuja was lowest (₦788,700)
- Beauty led all categories by revenue (₦915,400), narrowly ahead of Home & Kitchen (₦815,500); Apparel trailed all categories (₦448,000)
- Monthly sales were uneven — April was the strongest month (₦847,200) and February the weakest (₦330,900), suggesting a seasonal or promotional pattern worth investigating further

## Tools used

Microsoft Excel — TRIM/PROPER text cleaning, date-format reconciliation, Remove Duplicates, a native PivotTable with slicers, and bar/line charts.

## Note on the data

This dataset was generated for practice purposes and does not represent a real business. It was designed to include realistic data quality issues (mixed formats, duplicates, missing values, ambiguous dates) so the cleaning process could be demonstrated end to end.

## Author

**Elijah Bonny** — [linkedin.com/in/elijah-bonny](https://linkedin.com/in/elijah-bonny)
