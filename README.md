# E-commerce_sales
E-commerce sales capstone project
Below is a GitHub-ready README in Markdown. It is based on your uploaded Excel workbook, with verified totals, dashboard insights, methodology, and a transparent discussion of data-quality issues. The report distinguishes findings supported by the data from improvements that still need validation.


# E-Commerce Sales Analysis Using Microsoft Excel

## 1. Project Overview

This project analyzes an e-commerce dataset using Microsoft Excel to explore sales performance, customer demographics, product inventory, store distribution, and payment preferences.

The workbook brings together data cleaning, Excel formulas, lookup exercises, PivotTables, PivotCharts, and an interactive dashboard.

The analysis covers transactions recorded from **14 October 2023 to 13 October 2025**.

> Reporting note: Results describe the uploaded workbook and its cleaned data. Some cleaning formulas and dashboard configurations require additional validation; these limitations are documented below.

---

## 2. Business Objectives

The project addresses the following questions:

- What is the total recorded sales value?
- How many sales records and units are included?
- Which payment method contributes the most revenue?
- How does sales performance vary across the available years?
- How are customers distributed across gender and loyalty levels?
- How are stores distributed by region and store type?
- Which product categories hold the most recorded stock?
- What data-quality issues could affect interpretation?

---

## 3. Executive Summary

| Metric | Result |
|---|---:|
| Total recorded sales value | $2,175,257.98 |
| Sales records | 2,000 |
| Total units sold | 4,958 |
| Average sales value per record | $1,087.63 |
| Average units per sales record | 2.48 |
| Average recorded discount | 14.81% |
| Customer dimension records | 500 |
| Product dimension records | 100 |
| Store dimension records | 20 |
| Payment methods | 3 |
| Product categories | 5 |
| Date coverage | 14 Oct 2023–13 Oct 2025 |

### Main Findings

1. PayPal generated the highest recorded revenue: $756,714.87.
2. Payment-method revenue contributions were relatively balanced.
3. Platinum was the largest customer loyalty group, with 149 customers.
4. Female customers accounted for 52.4% of customer dimension records.
5. Sports held the highest recorded stock total.
6. Electronics had the highest average product cost.
7. East, North, and South each contained six stores; West contained two.
8. Annual revenue totals should not be interpreted as like-for-like growth because 2023 and 2025 have incomplete calendar-year coverage.

### Important Qualification

Customer counts are not customer sales, product stock is not product revenue, and store counts are not regional sales performance.

These measures should remain clearly separated throughout the dashboard and report.

---

## 4. Tools and Techniques

### Primary Tool

- Microsoft Excel

### Techniques Demonstrated

- Data profiling and cleaning
- Missing-value handling
- Identifier standardization
- Text case normalization
- Structured table references
- Logical and statistical formulas
- VLOOKUP
- PivotTables
- PivotCharts
- Date grouping
- Slicers
- Dashboard development

### Workbook Organization

The workbook contains separate sheets for source data, cleaned data, calculations, lookups, and presentation.

| Worksheet | Purpose |
|---|---|
| `Readme` | Workbook introduction |
| `Data clean_log` | Documented data-quality issues and actions |
| `Customer_raw` | Original customer data |
| `Customer_Dim_Clean` | Customer data used for analysis |
| `Product_raw` | Original product data |
| `Product_Dim_Clean` | Product data and helper columns |
| `Store_raw` | Original store data |
| `Store_Dim_Clean` | Store data used for analysis |
| `Sales_raw` | Original sales records |
| `Sales_Fact_Clean` | Sales data and helper columns |
| `Descriptive statistics and exce` | Statistical summaries and Excel calculations |
| `Pivot tables ` | Aggregated analysis |
| `Vlookup_Tables` | Lookup exercises |
| `Dashboard` | Visual reporting interface |

> Worksheet names above match the uploaded workbook, including shortened names.

---

## 5. Dataset Structure

The workbook contains one sales fact table and three supporting dimension tables.

### 5.1 Customer Dimension

| Field | Description |
|---|---|
| Customer_ID | Customer identifier |
| Name | Customer name |
| Age | Customer age |
| Gender | Recorded gender |
| City | Customer city |
| State | Customer state |
| Country | Customer country |
| Loyalty_Level | Customer loyalty category |

Customer dimension size: 500 records.

### 5.2 Product Dimension

| Field | Description |
|---|---|
| Product_ID | Product identifier |
| Product_Name | Product description |
| Category | Main product category |
| Sub_Category | Product subcategory |
| Brand | Product brand |
| Cost | Recorded product cost |
| Stock | Recorded inventory quantity |

Product dimension size: 100 records.

### 5.3 Store Dimension

| Field | Description |
|---|---|
| Store_ID | Store identifier |
| Store_Name | Store name |
| Region | Store region |
| City | Store city |
| Store_Type | Flagship, Online, or Outlet |

Store dimension size: 20 records.

### 5.4 Sales Fact Table

| Field | Description |
|---|---|
| Sales_ID | Sales record identifier |
| Order_Date | Recorded transaction date |
| Customer_ID | Customer reference |
| Product_ID | Product reference |
| Store_ID | Store reference |
| Quantity | Units recorded in the sale |
| Unit price | Selling price per unit |
| Discount | Discount recorded as a decimal |
| Payment_Type | Payment method |
| Total_Amount | Recorded sales amount |

Sales fact table size: 2,000 records.

### Intended Relationships

```text
Customer_Dim_Clean.Customer_ID ──< Sales_Fact_Clean.Customer_ID

Product_Dim_Clean.Product_ID ────< Sales_Fact_Clean.Product_ID

Store_Dim_Clean.Store_ID ────────< Sales_Fact_Clean.Store_ID
```

These keys provide the basis for transaction-level analysis by customer, product, and store.

> The presence of these keys does not by itself confirm that Excel Data Model relationships or shared slicer connections have been configured.

---

## 6. Data Preparation and Cleaning

The workbook preserves raw worksheets and includes a cleaning log for auditability.

### 6.1 Customer Data

The cleaning log records the following issues:

- Four customer identifiers in an inconsistent format.
- Three names with leading or trailing spaces.
- Twenty-nine names containing titles or suffixes.
- Four names with inconsistent capitalization.
- Five inconsistent country values.
- Four missing loyalty values.

Missing loyalty values were filled with `Platinum`.

### Validation Note

The cleaned customer sheet still contains examples of:

- Leading or trailing spaces in the original `Name` column.
- Lowercase names in the original `Name` column.
- Titles such as `Dr.` in some names.
- A helper column that applies proper case but does not necessarily remove spaces or titles.

Therefore, name standardization should be described as partially implemented rather than fully completed.

The cleaning log also states that the constant `Country` column was removed, but this column is still present in the cleaned worksheet.

### 6.2 Product Data

The cleaning log records:

- Two inconsistent product identifiers.
- Six missing cost values.
- Three missing stock values.

The cleaned product data contains replacement values, including fractional stock values.

### Validation Note

The helper column labeled `Stock imputed` repeats cost values in visible records because the formula’s nonblank branch references `Cost` rather than `Stock`.

This helper column should not be used for inventory analysis until corrected.

The cleaning log describes category-based mean imputation, but the implementation should be reconciled with that description before claiming that category-specific imputation was consistently applied.

### 6.3 Sales Data

The cleaning log records:

| Issue | Affected records or cells |
|---|---:|
| Customer-key mismatches | 22 |
| Product-key mismatches | 39 |
| Missing quantity | 5 |
| Missing unit price | 3 |
| Missing discount | 4 |
| Missing total amount | 4 |

The log documents identifier standardization and a combination of formula-based reconstruction and imputation.

The ten primary columns in `Sales_Fact_Clean` contain no missing values in the uploaded workbook.

> A complete table is not necessarily an error-free table. Reconstructed and imputed values should remain identifiable for review.

### 6.4 Recommended Cleaning Controls

Before treating the workbook as a production report:

1. Reconcile each cleaning-log statement with the actual worksheet contents.
2. Correct the stock helper formula.
3. Standardize names in one final reporting column.
4. Preserve flags identifying imputed values.
5. Validate all primary and foreign keys.
6. Reconcile transaction totals against their component fields.
7. Refresh PivotTables after corrections.

---

## 7. KPI Definitions

### Total Recorded Sales Value

```text
Total Recorded Sales Value = SUM(Total_Amount)
```

Result:

```text
$2,175,257.98
```

### Sales Record Count

```text
Sales Record Count = COUNT of Sales_ID
```

Result:

```text
2,000
```

This is a count of sales records. It should only be labeled an order count if the dataset’s business definition confirms that each row represents one complete order.

### Units Sold

```text
Units Sold = SUM(Quantity)
```

Result:

```text
4,958
```

### Average Sales Value per Record

```text
Average Sales Value per Record =
Total Recorded Sales Value / Sales Record Count
```

Result:

```text
$1,087.63
```

### Average Units per Sales Record

```text
Average Units per Sales Record =
Units Sold / Sales Record Count
```

Result:

```text
2.48
```

### Average Recorded Discount

```text
Average Recorded Discount = AVERAGE(Discount)
```

Result:

```text
14.81%
```

This is an unweighted average of recorded discount rates. It is not the same as a revenue-weighted effective discount.

### Expected Transaction Amount

Where the business rule is applicable:

```text
Expected Total Amount =
Quantity × Unit Price × (1 − Discount)
```

This calculation can be used to identify records requiring reconciliation.

---

## 8. Sales Performance Analysis

### 8.1 Revenue by Available Calendar Year

| Year | Recorded sales value | Units sold | Coverage qualification |
|---|---:|---:|---|
| 2023 | $224,165.46 | 516 | Begins 14 October |
| 2024 | $1,093,872.09 | 2,486 | Full calendar year within the dataset period |
| 2025 | $857,220.42 | 1,956 | Ends 13 October |
| Total | $2,175,257.98 | 4,958 | Entire available period |

> Displayed annual amounts are rounded independently, so their displayed sum can differ from the rounded grand total by one cent.

### Interpretation

2024 has the highest recorded annual revenue in the available data.

However, the years have different reporting windows:

- 2023 includes only part of the fourth quarter.
- 2024 includes the entire calendar year.
- 2025 ends before the calendar year is complete.

Consequently, the lower 2025 total does not establish a year-over-year business decline.

### Better Comparison Approach

For a fairer performance comparison:

- Compare matching date windows.
- Compare complete months or complete quarters.
- Separate partial periods from complete periods.
- Use comparable-period revenue growth rather than raw annual totals.

---

## 9. Payment Method Analysis

| Payment method | Recorded sales value | Sales records | Units sold | Approximate revenue share |
|---|---:|---:|---:|---:|
| PayPal | $756,714.87 | 695 | 1,696 | 34.8% |
| Credit Card | $714,842.56 | 645 | 1,618 | 32.9% |
| COD | $703,700.55 | 660 | 1,644 | 32.4% |
| Total | $2,175,257.98 | 2,000 | 4,958 | 100% |

> Individually rounded shares may not sum to exactly 100%.

### Key Insights

- PayPal leads in recorded revenue, sales-record count, and units sold.
- Credit Card generates more revenue than COD despite having fewer sales records.
- All three methods make substantial contributions.
- No payment method accounts for a majority of revenue.

### Business Interpretation

The data supports maintaining visibility across all three payment channels.

However, the highest-revenue payment method is not necessarily the most profitable. Payment fees, failed transactions, refunds, and COD collection costs are not included in this comparison.

---

## 10. Customer Analysis

### 10.1 Gender Distribution

| Recorded gender | Customers | Share |
|---|---:|---:|
| Female | 262 | 52.4% |
| Male | 238 | 47.6% |
| Total | 500 | 100% |

### Interpretation

The customer dimension is relatively balanced across the two recorded gender categories, with a small female majority.

This describes customer composition, not revenue contribution.

### 10.2 Loyalty-Level Distribution

| Loyalty level | Customers | Share |
|---|---:|---:|
| Platinum | 149 | 29.8% |
| Bronze | 125 | 25.0% |
| Gold | 122 | 24.4% |
| Silver | 104 | 20.8% |
| Total | 500 | 100% |

### Key Insights

- Platinum is the largest recorded loyalty group.
- Silver is the smallest recorded loyalty group.
- Bronze and Gold have similar customer counts.

### Qualification

Four missing loyalty values were filled with `Platinum`, so the final distribution includes imputed classifications.

The largest loyalty group should not automatically be described as the most valuable group. That requires transaction-level revenue, purchase frequency, and margin analysis.

### Geographic Scope

All 500 customer records are classified as USA in the cleaned customer dimension.

As a result, `Country` does not distinguish customer segments within this dataset.

### Dashboard Naming Improvement

The chart labeled `Customer Sales Analysis` is based on customer counts rather than sales amounts.

A more accurate title would be:

```text
Customer Distribution by Gender and Loyalty Level
```

---

## 11. Product and Inventory Analysis

| Category | Recorded stock total | Average recorded product cost |
|---|---:|---:|
| Sports | 7,040.49 | $211.80 |
| Electronics | 5,767.49 | $316.56 |
| Clothing | 4,931.00 | $263.78 |
| Beauty | 4,816.00 | $250.90 |
| Home | 4,794.49 | $235.72 |

### Key Insights

- Sports has the highest recorded stock total.
- Electronics has the highest average product cost.
- Home has the lowest recorded stock total among the five categories.
- Stock quantity and product cost describe different aspects of the catalog.

### Interpretation Limits

High inventory does not establish strong demand.

A category with high stock could reflect:

- A broad product range.
- Recent replenishment.
- Slow inventory movement.
- A deliberate availability strategy.

These explanations cannot be distinguished from the stock snapshot alone.

### Inventory Data-Quality Issue

Fractional values appear in the stock totals because missing stock values were replaced with noninteger estimates.

If stock represents indivisible units, those values require a documented rounding or alternative imputation policy.

### Cost Versus Inventory Value

The sum of product `Cost` values is not total inventory value.

If `Cost` represents unit cost, inventory value should instead be calculated at the product level:

```text
Inventory Value = SUMPRODUCT(Unit Cost, Stock)
```

This calculation should only be finalized after correcting stock-related issues.

### Dashboard Naming Improvement

A chart combining stock and cost should be labeled:

```text
Product Inventory and Cost Analysis
```

It should not be presented as product sales performance unless transaction-level sales measures are included.

---

## 12. Store Distribution Analysis

### 12.1 Stores by Region

| Region | Stores | Share |
|---|---:|---:|
| East | 6 | 30% |
| North | 6 | 30% |
| South | 6 | 30% |
| West | 2 | 10% |
| Total | 20 | 100% |

### Key Insight

East, North, and South have equal store counts. West has the smallest store footprint.

This describes store distribution, not regional revenue or productivity.

### Dashboard Reconciliation Note

Some dashboard screenshots display regional percentages that do not match the complete 20-store distribution above.

Before publishing the dashboard:

- Clear all filters and slicers.
- Refresh the relevant PivotTable.
- Check the source range.
- Confirm the chart uses the intended measure.
- Verify whether the displayed percentages represent a filtered subset.

### 12.2 Stores by Type

| Store type | Stores | Share |
|---|---:|---:|
| Flagship | 8 | 40% |
| Online | 8 | 40% |
| Outlet | 4 | 20% |
| Total | 20 | 100% |

### Interpretation

Flagship and Online represent equal shares of the store dimension.

These counts do not indicate how much revenue each store type generates.

A sales-based store comparison requires joining sales records to store attributes using `Store_ID`.

---

## 13. Excel Formula Examples

The formulas below illustrate recommended patterns.

They are not a claim that every formula has already been implemented correctly in the workbook.

### 13.1 Normalize Name Spacing and Capitalization

```excel
=PROPER(TRIM([@Name]))
```

This handles extra spaces and letter case.

It does not reliably remove every title or suffix, and proper case may require manual review for some names.

### 13.2 Fill Missing Cost with an Overall Mean

```excel
=ROUND(IF(ISBLANK([@Cost]),AVERAGE(Table2_P[Cost]),[@Cost]),2)
```

Use category-based imputation instead when that is the documented cleaning policy.

### 13.3 Correct the Stock Helper Reference

```excel
=IF(ISBLANK([@Stock]),AVERAGE(Table2_P[Stock]),[@Stock])
```

If stock must be a whole number:

```excel
=ROUND(IF(ISBLANK([@Stock]),AVERAGE(Table2_P[Stock]),[@Stock]),0)
```

The choice to round should be documented rather than applied silently.

### 13.4 Recalculate a Transaction Amount

```excel
=ROUND([@Quantity]*[@[Unit price]]*(1-[@Discount]),2)
```

### 13.5 Retrieve Store Attributes

```excel
=VLOOKUP([@[Store_ID]],Store_Dim_Clean!A:E,2,FALSE)
```

This example retrieves the store name using an exact match.

### 13.6 Validate a Customer Key

```excel
=IF(COUNTIF(Customer_Dim_Clean!A:A,[@Customer_ID])=1,"Valid","Review")
```

### 13.7 Flag Missing Loyalty Values

```excel
=IF(ISBLANK([@Loyalty_Level]),"Missing","Observed")
```

Maintaining an imputation flag makes later reporting more transparent.

> Formula examples assume the relevant structured table and column names exist. Adjust references to match the workbook.

---

## 14. Lookup Tables and Data Integration

The workbook includes VLOOKUP exercises that retrieve attributes from source worksheets.

These exercises demonstrate exact-match lookup logic.

However, placing multiple dimension tables side by side does not create a transaction-level analytical join.

For example:

- A sales row can reference `CUST1409`.
- A customer table displayed alongside it can begin with `CUST1000`.
- Their physical position on the same row does not mean they refer to the same customer.

### Correct Integration Principle

Each sale must retrieve attributes using its own identifiers:

```text
Sales.Customer_ID → Matching customer attributes

Sales.Product_ID → Matching product attributes

Sales.Store_ID → Matching store attributes
```

This distinction matters when calculating:

- Revenue by loyalty level.
- Sales by customer gender.
- Revenue by product category.
- Sales by region.
- Revenue by store type.

---

## 15. Dashboard Overview

The dashboard presents several analytical views.

| View | Intended purpose |
|---|---|
| Product chart | Compare category/subcategory inventory and cost |
| Customer chart | Compare customer counts by gender and loyalty |
| Region/store chart | Display store distribution |
| Payment chart | Compare recorded revenue by payment method |
| Time-series chart | Explore recorded sales across available dates |

### Filters Shown in the Workbook Screenshots

- Category
- Loyalty level
- Region
- Payment type
- Year

### Recommended Dashboard Improvements

1. Add KPI cards for revenue, sales records, units, and average sales value.
2. Rename charts so titles match the actual measures.
3. Keep monetary values, quantities, and percentages on appropriately labeled axes.
4. Avoid placing unrelated measures on the same scale.
5. Replace 3D pie charts with simpler charts for clearer comparisons.
6. Use consistent category colors.
7. Hide chart field buttons in the final presentation where appropriate.
8. Clearly indicate active filters.
9. Validate slicer report connections.
10. Add a visible reporting-period label.

### Date-Chart Improvement

Use one consistent date grain at a time, such as:

- Month
- Quarter
- Year

Do not interpret a chart that mixes expanded monthly detail with collapsed quarterly totals as a uniform time series.

### Discount-Chart Improvement

Plot discount on a percentage scale or in a separate chart.

A discount rate near 15% is not meaningfully readable on the same axis as revenue values in the hundreds of thousands.

---

## 16. Recommended Next Analyses

### 16.1 Comparable-Period Sales

Compare the same date windows across years before drawing growth conclusions.

### 16.2 Customer Value

Join sales to customer records to calculate:

- Revenue per customer.
- Sales-record frequency.
- Revenue by loyalty level.
- Repeat-purchase behavior.

### 16.3 Product Performance

Join sales to products to calculate:

- Revenue by category.
- Units sold by product.
- Revenue concentration.
- Inventory movement indicators.

### 16.4 Store Productivity

Join sales to stores to calculate:

- Revenue by region.
- Revenue per store.
- Revenue by store type.
- Units sold per store.

### 16.5 Payment Economics

Add fee, refund, and collection-cost data before recommending a payment-method strategy.

### 16.6 Discount Effectiveness

Compare discount levels with transaction values and units sold.

Any observed relationship should be treated as descriptive, not proof that discounts caused higher sales.

---

## 17. Data Quality and Analytical Limitations

| Limitation | Why it matters |
|---|---|
| Partial calendar years | Raw annual totals are not directly comparable |
| Imputed loyalty values | Customer segmentation includes estimated classifications |
| Fractional stock values | Inventory may require an integer-based business rule |
| Incorrect stock helper reference | Helper results can misstate inventory |
| Name-cleaning inconsistencies | Cleaning documentation overstates some completed work |
| Cleaning-log discrepancies | Methods described do not always match visible outputs |
| Filter-dependent chart percentages | Dashboard values may represent a subset |
| Mixed measures in charts | Quantities, money, and percentages can be misleading together |
| Side-by-side lookup tables | Row alignment does not establish a valid join |
| No documented order-grain confirmation | Sales-record count may not equal unique order count |
| No payment-cost information | Revenue leadership does not establish profitability |
| No historical inventory series | Inventory turnover cannot be inferred from a stock snapshot |

### Revenue Qualification

The main reported revenue total uses the primary `Total_Amount` column in `Sales_Fact_Clean`.

Helper columns may contain rounded or differently imputed values. Use one validated measure consistently across calculations and PivotTables.

### Currency Qualification

Amounts are displayed using the workbook’s dollar formatting. Confirm the intended currency code before publishing the analysis as a formal financial report.

---

## 18. How to Explore the Project

1. Download `EC-Sales.xlsx`.
2. Open the workbook in Microsoft Excel.
3. Review the `Readme` and `Data clean_log` worksheets.
4. Compare raw and cleaned data.
5. Review helper formulas before relying on their outputs.
6. Open the `Pivot tables ` worksheet.
7. Explore the `Dashboard`.
8. Clear filters when reconciling results with this report.
9. Refresh PivotTables after making changes.

> Preserve an original copy of the workbook before editing formulas, source ranges, or dashboard connections.

---

## 19. Suggested Repository Structure

The following is a suggested layout, not a statement that these folders already exist:

```text
ecommerce-sales-analysis/
│
├── README.md
├── EC-Sales.xlsx
│
├── screenshots/
│   ├── dashboard.png
│   ├── data-cleaning.png
│   ├── customer-analysis.png
│   ├── product-analysis.png
│   ├── payment-analysis.png
│   └── store-analysis.png
│
└── reports/
    └── analysis-screenshots.pdf
```

Only add screenshot links after uploading the corresponding files to the repository.

---

## 20. Skills Demonstrated

- Excel-based data preparation
- Data-quality documentation
- Missing-value handling
- Lookup formulas
- Structured table references
- PivotTable aggregation
- Customer segmentation summaries
- Inventory and cost comparisons
- Payment-method analysis
- Interactive dashboard design
- Analytical validation
- Business reporting

---

## 21. Conclusion

The workbook records approximately $2.18 million in sales across 2,000 sales records and 4,958 units during the available reporting period.

PayPal contributes the largest share of recorded revenue, while all three payment methods remain important. Platinum is the largest customer loyalty group, Sports holds the highest recorded stock, and Electronics has the highest average product cost.

The project demonstrates an end-to-end Excel reporting workflow, from source-data review to dashboard presentation.

Its most important next step is validation: correct the stock helper formula, reconcile the cleaning log with actual outputs, standardize reporting measures, and verify dashboard filters and relationships.

With those improvements, the workbook can provide a clearer and more dependable foundation for customer, product, store, and sales analysis.
````
