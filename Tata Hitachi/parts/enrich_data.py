import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')

base_dir = r"C:\Users\ayush\Downloads\internship\parts"
clean_dir = os.path.join(base_dir, "clean_data")
os.makedirs(clean_dir, exist_ok=True)

print("Starting Data Enrichment and Clean Process...")

# ============================================================
# 1. ENRICH MOP TRACKER
# ============================================================
mop_file = os.path.join(base_dir, "MOP tracker sheet upto March 2026.xlsx")
mop_df = pd.read_excel(mop_file)

# Remove rows where 'Sr. No.' is empty
mop_df = mop_df.dropna(subset=['Sr. No.'])
mop_df.columns = mop_df.columns.str.strip().str.replace('\n', ' ')

# Clean monthly sales columns to numeric
monthly_cols = ["April'25", "May'25", "June'25", "July'25", "Aug'25", "Sep'25", "Oct'25", "Nov'25", "Dec'25", "Jan'26", "Feb'26", "Mar'26"]
for c in monthly_cols:
    if c in mop_df.columns:
        mop_df[c] = mop_df[c].astype(str).str.strip().str.replace(',', '').str.replace(' ', '')
        mop_df[c] = pd.to_numeric(mop_df[c], errors='coerce').fillna(0.0)

# Calculate Actual Sales Revenue
mop_df['Actual Sales Revenue'] = mop_df[monthly_cols].sum(axis=1)

# Clean ORDER QTY column to be numeric, handling cases like "2+1"
if 'ORDER QTY' in mop_df.columns:
    mop_df['ORDER QTY'] = mop_df['ORDER QTY'].astype(str).str.strip().str.replace(' ', '')
    mop_df['ORDER QTY'] = mop_df['ORDER QTY'].replace('2+1', '3')
    mop_df['ORDER QTY'] = pd.to_numeric(mop_df['ORDER QTY'], errors='coerce')

# Determine if the order was Won
has_invoice = mop_df['INVOICE NO.'].notna() & (mop_df['INVOICE NO.'].astype(str).str.strip() != '')
has_order = mop_df['ORDER QTY'].notna() & (mop_df['ORDER QTY'].astype(str).str.strip() != '')
has_revenue = mop_df['Actual Sales Revenue'] > 0.0
mop_df['Is Won'] = np.where(has_invoice | has_order | has_revenue, 1, 0)

# 1. Lost Revenue (Quotation price - actual sales)
mop_df['Lost Revenue'] = np.maximum(0.0, mop_df['QUOTATION- Total Price'].fillna(0.0) - mop_df['Actual Sales Revenue'])

# 2. Discount Given
disc_price_col = 'Price After Discount  (Quoted Price without GST)               Quotation Amount'
if disc_price_col in mop_df.columns:
    mop_df['Discount Given'] = np.maximum(0.0, mop_df['QUOTATION- Total Price'].fillna(0.0) - mop_df[disc_price_col].fillna(mop_df['QUOTATION- Total Price']))
else:
    mop_df['Discount Given'] = 0.0

# 3. Response Days (Report to Quote) & Cycle Days (Quote to Invoice)
# Use temporary variables for date calculations so we don't convert the original columns to datetime and cause type mismatches in Power BI
report_date_dt = pd.to_datetime(mop_df['MOP REPORT DATE'], errors='coerce')
quote_date_dt = pd.to_datetime(mop_df['MOP QUOTATION DATE'], errors='coerce')
invoice_date_dt = pd.to_datetime(mop_df['INVOICE DATE'], errors='coerce')

mop_df['Response Days (Report to Quote)'] = (quote_date_dt - report_date_dt).dt.days
mop_df['Response Days (Report to Quote)'] = np.where(mop_df['Response Days (Report to Quote)'] >= 0, mop_df['Response Days (Report to Quote)'], np.nan)

mop_df['Cycle Days (Quote to Invoice)'] = (invoice_date_dt - quote_date_dt).dt.days
mop_df['Cycle Days (Quote to Invoice)'] = np.where(mop_df['Cycle Days (Quote to Invoice)'] >= 0, mop_df['Cycle Days (Quote to Invoice)'], np.nan)

# Restore the old 'Lost Quotation Value' column so that existing Power BI visuals/steps referencing it do not break
mop_df['Lost Quotation Value'] = np.where(mop_df['Is Won'] == 0, mop_df['QUOTATION- Total Price'].fillna(0.0), 0.0)

mop_out = os.path.join(clean_dir, "MOP_Tracker_Clean.csv")
mop_df.to_csv(mop_out, index=False)
print(f"MOP Tracker enriched: {mop_df.shape[0]} rows, {mop_df.shape[1]} columns -> {mop_out}")

# ============================================================
# 2. ENRICH MACHINE POPULATION
# ============================================================
pop_file = os.path.join(base_dir, "Final Machine Population 26-27(09-04-26).xlsx")
pop_df = pd.read_excel(pop_file)
pop_df.columns = pop_df.columns.str.strip().str.replace('\n', ' ')

# Calculate Machine Age and Age Buckets
# Use a temporary datetime series so we don't alter the original DOC column format
doc_dt = pd.to_datetime(pop_df['DOC'], errors='coerce')
pop_df['Machine Age Years'] = 2026 - doc_dt.dt.year

def get_age_bucket(age):
    if pd.isna(age):
        return "Unknown"
    elif age <= 3:
        return "0-3 Years"
    elif age <= 7:
        return "4-7 Years"
    elif age <= 12:
        return "8-12 Years"
    else:
        return "13+ Years"

pop_df['Age Bucket'] = pop_df['Machine Age Years'].apply(get_age_bucket)

pop_out = os.path.join(clean_dir, "Machine_Population_Clean.csv")
pop_df.to_csv(pop_out, index=False)
print(f"Machine Population enriched: {pop_df.shape[0]} rows, {pop_df.shape[1]} columns -> {pop_out}")

# ============================================================
# 3. CLEAN FDV LOG BOOK - MERGE ALL 12 MONTHLY SHEETS (FY 2025-26)
# ============================================================
try:
    fdv_file = os.path.join(base_dir, "FDV LOG BOOK UPTO Mar-26 (2025-26).xls")
    xl = pd.ExcelFile(fdv_file)

    # All 12 monthly sheets for fiscal year April 2025 - March 2026
    sheets_fy2526 = [
        'APR-25', 'MAY-25', 'JUNE-25', 'JULY-25', 'AUG-25', 'SEPT-25',
        'OCT-25', 'NOV-25', 'DEC-25', 'JAN-26', 'FEB 26', 'MAR 26'
    ]

    # Standard column names for all sheets
    standard_cols = [
        'Date', 'From', 'To', 'Starting KM', 'Closing KM',
        'Location of site', 'NAME OF THE CUSTOMER', 'Machine Sl.no',
        'Job card Number', 'Type of Machine', 'Nature of Job',
        'Revenue-Service(Rs)', 'OFFER NO/DATE', 'OFFER AMOUNT',
        'Revenue-Parts(Rs)', 'Remarks'
    ]

    all_dfs = []
    for sheet_name in sheets_fy2526:
        if sheet_name in xl.sheet_names:
            # Each sheet has 6 metadata rows at the top, then the header row
            df = xl.parse(sheet_name, skiprows=6)
            df.columns = standard_cols  # Force standard column names
            df = df.dropna(subset=['Date'], how='all')  # Remove rows without a date
            df['Source Month'] = sheet_name  # Track which month this row came from
            all_dfs.append(df)
            print(f"   Loaded sheet '{sheet_name}': {df.shape[0]} rows")
        else:
            print(f"   [WARN] Sheet '{sheet_name}' not found, skipping.")

    if all_dfs:
        fdv_combined = pd.concat(all_dfs, ignore_index=True)

        # Parse dates and remove invalid rows
        fdv_combined['Date'] = pd.to_datetime(fdv_combined['Date'], errors='coerce')
        fdv_combined = fdv_combined.dropna(subset=['Date'])
        fdv_combined = fdv_combined.sort_values(by='Date').reset_index(drop=True)

        # Remove footer/summary rows (like "THCM Branch Manager" etc.)
        fdv_combined = fdv_combined[fdv_combined['From'].notna()]
        # Remove rows where 'From' contains manager/summary text
        summary_keywords = ['THCM', 'Branch Manager', 'DSE', 'DPSE']
        mask = ~fdv_combined['From'].astype(str).str.contains('|'.join(summary_keywords), case=False, na=False)
        fdv_combined = fdv_combined[mask].reset_index(drop=True)

        # Calculate Trip Distance (km)
        fdv_combined['Starting KM'] = pd.to_numeric(fdv_combined['Starting KM'], errors='coerce')
        fdv_combined['Closing KM'] = pd.to_numeric(fdv_combined['Closing KM'], errors='coerce')
        fdv_combined['Trip Distance KM'] = fdv_combined['Closing KM'] - fdv_combined['Starting KM']
        fdv_combined['Trip Distance KM'] = np.where(fdv_combined['Trip Distance KM'] >= 0, fdv_combined['Trip Distance KM'], np.nan)

        # Clean Revenue columns
        fdv_combined['Revenue-Service(Rs)'] = pd.to_numeric(fdv_combined['Revenue-Service(Rs)'], errors='coerce').fillna(0)
        fdv_combined['Revenue-Parts(Rs)'] = pd.to_numeric(fdv_combined['Revenue-Parts(Rs)'], errors='coerce').fillna(0)
        fdv_combined['OFFER AMOUNT'] = pd.to_numeric(fdv_combined['OFFER AMOUNT'], errors='coerce').fillna(0)

        fdv_out = os.path.join(clean_dir, "FDV_Log_Book_Clean.csv")
        fdv_combined.to_csv(fdv_out, index=False)
        print(f"FDV Log Book merged: {fdv_combined.shape[0]} rows, {fdv_combined.shape[1]} columns -> {fdv_out}")
    else:
        print("[ERROR] No FDV sheets could be loaded!")

except Exception as e:
    print(f"[ERROR] Error cleaning FDV Log Book: {e}")
    import traceback
    traceback.print_exc()

print("\nData enrichment process complete! All 3 CSV files are ready for Power BI.")
