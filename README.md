# Sales Data Cleaning and Quality Verification

## Project Overview

This project checks sales data quality, cleans the dataset, and verifies the cleaned output before analysis.

The objective is to prepare consistent data while highlighting issues that require client confirmation or special treatment during analysis.

**The dataset included in this repository is a representative sample. It is not the original client dataset.**

## Workflow

1. Verify the original dataset and identify data quality issues.
2. Clean the data according to the agreed scope.
3. Verify the cleaned dataset.
4. Share the cleaned data and verification report, highlighting exceptions and decisions required before analysis.

## Cleaning Requirements

The agreed cleaning scope included:

- Standardizing column names.
- Standardizing date formats and text values, such as names.
- Representing missing values consistently as `NA`.

## Client Requirements for Analysis

### Missing Customer IDs

- Retain these transactions for overall sales analysis.
- Exclude them from customer-specific metrics.
- State that customer findings cover identified customers only.
- Report the percentage of transactions without a Customer ID.

### Missing Product Descriptions

- Use `StockCode` for product grouping.
- Check whether each stock code identifies a product consistently.
- Flag stock codes associated with conflicting descriptions for review.

### Zero or Near-Zero Prices

- Preserve the recorded prices unless the client confirms a correction.
- Flag zero prices separately.
- Define near-zero prices as values greater than `0` and less than `0.1`, in the dataset’s price currency.
- Provide an exception summary containing recorded prices, transaction counts, quantities, and example invoice numbers.

These records require review because their correctness cannot be confirmed from the supplied files alone.

### Negative Prices

Report records identified as bad-debt adjustments separately from product sales. Other negative-price records require review before classification.

### Zero Quantities

Flag records with a quantity of `0` for review.

## Deliverables and Results

The cleaned dataset and verification report were shared with the client according to the agreed scope.

The report highlighted remaining exceptions and areas requiring client decisions before analysis. Completing the cleaning process does not mean that every business-related exception has been resolved.

## Tools

- Python
- pandas
- pathlib
- re
