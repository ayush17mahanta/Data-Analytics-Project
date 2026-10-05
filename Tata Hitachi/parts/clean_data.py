import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')

# Set working directory
base_dir = r"C:\Users\ayush\Downloads\internship\parts"
clean_dir = os.path.join(base_dir, "clean_data")
os.makedirs(clean_dir, exist_ok=True)

print("Starting data cleaning process...")

# 1. Clean MOP Tracker
try:
    mop_file = os.path.join(base_dir, "MOP tracker sheet upto March 2026.xlsx")
    mop_df = pd.read_excel(mop_file)
    
    # Remove rows where 'Sr. No.' is empty to get rid of completely blank rows at bottom
    mop_df = mop_df.dropna(subset=['Sr. No.'])
    
    # Clean column names
    mop_df.columns = mop_df.columns.str.strip().str.replace('\n', ' ')
    
    # Clean monthly sales columns to numeric
    monthly_cols = ["April'25", "May'25", "June'25", "July'25", "Aug'25", "Sep'25", "Oct'25", "Nov'25", "Dec'25", "Jan'26", "Feb'26", "Mar'26"]
    
    for c in monthly_cols:
        if c in mop_df.columns:
            mop_df[c] = mop_df[c].astype(str).str.strip().str.replace(',', '').str.replace(' ', '')
            mop_df[c] = pd.to_numeric(mop_df[c], errors='coerce').fillna(0.0)
            
    # Calculate Actual Sales Revenue
    mop_df['Actual Sales Revenue'] = mop_df[monthly_cols].sum(axis=1)
    
    # Determine if the order was Won
    # An order is won if it has an Invoice Number, Order Qty, or Actual Sales Revenue > 0
    has_invoice = mop_df['INVOICE NO.'].notna() & (mop_df['INVOICE NO.'].astype(str).str.strip() != '')
    has_order = mop_df['ORDER QTY'].notna() & (mop_df['ORDER QTY'].astype(str).str.strip() != '')
    has_revenue = mop_df['Actual Sales Revenue'] > 0.0
    
    mop_df['Is Won'] = np.where(has_invoice | has_order | has_revenue, 1, 0)
    
    # Calculate Lost Quotation Value
    mop_df['Lost Quotation Value'] = np.where(mop_df['Is Won'] == 0, mop_df['QUOTATION- Total Price'], 0.0)
    
    mop_out = os.path.join(clean_dir, "MOP_Tracker_Clean.csv")
    mop_df.to_csv(mop_out, index=False)
    print(f"Successfully cleaned MOP Tracker and added computed columns -> {mop_out}")
except Exception as e:
    print(f"Error cleaning MOP Tracker: {e}")

# 2. Clean Machine Population
try:
    pop_file = os.path.join(base_dir, "Final Machine Population 26-27(09-04-26).xlsx")
    pop_df = pd.read_excel(pop_file)
    pop_df.columns = pop_df.columns.str.strip().str.replace('\n', ' ')
    
    pop_out = os.path.join(clean_dir, "Machine_Population_Clean.csv")
    pop_df.to_csv(pop_out, index=False)
    print(f"Successfully cleaned Machine Population -> {pop_out}")
except Exception as e:
    print(f"Error cleaning Machine Population: {e}")

# 3. Clean FDV Log Book
try:
    fdv_file = os.path.join(base_dir, "FDV LOG BOOK UPTO Mar-26 (2025-26).xls")
    fdv_df = pd.read_excel(fdv_file, skiprows=5)
    
    fdv_df = fdv_df.dropna(how='all')
    fdv_df.columns = fdv_df.columns.astype(str).str.strip().str.replace('\n', ' ')
    fdv_df = fdv_df.loc[:, ~fdv_df.columns.str.contains('^Unnamed')]
    
    fdv_out = os.path.join(clean_dir, "FDV_Log_Book_Clean.csv")
    fdv_df.to_csv(fdv_out, index=False)
    print(f"Successfully cleaned FDV Log Book -> {fdv_out}")
except Exception as e:
    print(f"Error cleaning FDV Log Book: {e}")

print("Data cleaning complete! Files are ready for Power BI.")
