import pandas as pd
import numpy as np

# Load MOP Tracker
df = pd.read_csv(r"C:\Users\ayush\Downloads\internship\parts\clean_data\MOP_Tracker_Clean.csv")

# Clean monthly columns
monthly_cols = [c for c in df.columns if "'" in c]
print("Monthly columns found:", monthly_cols)

for c in monthly_cols:
    # Convert to string first to strip, replace commas, then convert to numeric
    df[c] = df[c].astype(str).str.strip().str.replace(',', '').str.replace(' ', '')
    df[c] = pd.to_numeric(df[c], errors='coerce')

# Print info on cleaned monthly columns
print("\nCleaned monthly columns sum:")
print(df[monthly_cols].sum())

total_sales_sum = df[monthly_cols].sum().sum()
print("\nTotal Sales Sum (from monthly columns):", total_sales_sum)

# Let's count rows that have actual orders
has_invoice = df['INVOICE NO.'].notna()
has_order_qty = df['ORDER QTY'].notna()
has_monthly_sale = df[monthly_cols].notna().any(axis=1)

print("\nRow counts:")
print("Total rows:", len(df))
print("Rows with INVOICE NO.:", has_invoice.sum())
print("Rows with ORDER QTY:", has_order_qty.sum())
print("Rows with monthly sales values:", has_monthly_sale.sum())

# Let's see what is the sum of QUOTATION- Total Price
print("\nSum of QUOTATION- Total Price:", df['QUOTATION- Total Price'].sum())
print("Sum of Price After Discount:", df['Price After Discount  (Quoted Price without GST)               Quotation Amount'].sum())
print("Sum of QUOTATION Final Price:", df['QUOTATION Final Price'].sum())

# Let's check how many orders are won (where INVOICE NO is not null or monthly sale exists)
won_df = df[has_monthly_sale | has_invoice]
print("\nWon orders info:")
print("Number of won orders:", len(won_df))
print("Total revenue from won orders (sum of monthly sales):", won_df[monthly_cols].sum().sum())
print("Total quoted value for won orders (QUOTATION- Total Price):", won_df['QUOTATION- Total Price'].sum())
print("Total quoted value with GST for won orders (QUOTATION Final Price):", won_df['QUOTATION Final Price'].sum())

# Win rate by quantity of quotes
print("\nWin Rate (Quotes with sales / Total Quotes):", len(won_df) / len(df))

# Let's check if the user had some numeric issues with columns in Power BI
