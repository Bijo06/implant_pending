import pandas as pd
import numpy as np

source_file = r'C:\Users\ppc.controllers\Desktop\imp-pending\PPC(09022026) SAP IMPLANT HOURLY REPORT 19.02.2026.xlsx'
output_file = r'C:\Users\ppc.controllers\Desktop\imp-pending\IMPLANT PENDING.xlsx'

# 1. Read Source Data
print("Reading source file...")
df_raw = pd.read_excel(source_file, sheet_name='Sheet1')

# 2. Apply Filters to match the 157 total
df = df_raw.copy()

# JobCardNum not blank and doesn't start with M
df = df[df['JobCardNum'].astype(str).str.strip() != '']
df = df[~df['JobCardNum'].astype(str).str.startswith('M')]

# CustomerName exclusions
blocked_customers = ['SMILE DENTAL CARE', 'EMPIRE DENTAL CARE']
df = df[~df['CustomerName'].isin(blocked_customers)]

# ProductName exclusions
df = df[~df['ProductName'].str.contains('TRIAL', na=False)]

# FamilyName filter
df = df[df['FamilyName'].isin(['METAL', 'METAL FREE'])]

# LastScanningLocation exclusions
blocked_locations = ['FINALATTACHEMENTPACK', 'QA', 'WAITING FOR CONNECTED JOBCARD', 'BLANK', '-BLANK']
df = df[df['LastScanningLocation'].astype(str).str.strip() != '']
df = df[~df['LastScanningLocation'].isin(blocked_locations)]

# RequiredDate <= today (2026-05-14)
df['RequiredDate_Parsed'] = pd.to_datetime(df['RequiredDate'], errors='coerce')
today = pd.Timestamp('2026-05-14')
df = df[df['RequiredDate_Parsed'] <= today]

print(f"Filtered rows: {len(df)}")
print(f"Grand Total (TeethCount sum): {df['TeethCount'].sum()}")

# 3. Prepare Sheets
# Sheet1: The filtered data
# PENDING: subset of columns
pending_cols = [
    'JobCardNum', 'PriorityCode', 'StateCode', 'CustomerName', 
    'ProductName', 'FamilyName', 'TeethCount', 'RegDate', 
    'RequiredDate', 'LastScanningLocation', 'New Department'
]
df_pending = df[pending_cols].copy()

# 4. Create Pivot Table (Sheet3)
# Rows=New Department, Cols=RequiredDate_Parsed, Values=TeethCount
# We'll format the dates for the columns
df['DateStr'] = df['RequiredDate_Parsed'].dt.strftime('%d-%m-%Y')

pivot_df = df.pivot_table(
    index='New Department',
    columns='DateStr',
    values='TeethCount',
    aggfunc='sum',
    margins=True,
    margins_name='Grand Total'
)

# 5. Save to Excel
print(f"Saving to {output_file}...")
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Sheet1
    df.drop(columns=['RequiredDate_Parsed', 'DateStr']).to_excel(writer, sheet_name='Sheet1', index=False)
    
    # PENDING
    df_pending.to_excel(writer, sheet_name='PENDING', index=False)
    
    # Sheet3
    workbook = writer.book
    worksheet3 = workbook.add_worksheet('Sheet3')
    
    # Headers
    worksheet3.write('A2', 'Sum of TeethCount')
    worksheet3.write('B2', 'Column Labels')
    worksheet3.write(2, 0, 'Row Labels')
    
    # Date columns headers
    date_cols = [c for c in pivot_df.columns if c != 'Grand Total']
    date_cols.sort(key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
    
    for i, date in enumerate(date_cols):
        worksheet3.write(2, i + 1, date)
    worksheet3.write(2, len(date_cols) + 1, 'Grand Total')
    
    # Data rows
    row_labels = [r for r in pivot_df.index if r != 'Grand Total']
    row_labels.sort()
    
    for r_idx, label in enumerate(row_labels):
        worksheet3.write(r_idx + 3, 0, label)
        for c_idx, date in enumerate(date_cols):
            val = pivot_df.loc[label, date]
            if not pd.isna(val) and val != 0:
                worksheet3.write(r_idx + 3, c_idx + 1, val)
        # Total for row
        row_total = pivot_df.loc[label, 'Grand Total']
        worksheet3.write(r_idx + 3, len(date_cols) + 1, row_total)
        
    # Grand Total row
    gt_idx = len(row_labels) + 3
    worksheet3.write(gt_idx, 0, 'Grand Total')
    for c_idx, date in enumerate(date_cols):
        col_total = pivot_df.loc['Grand Total', date]
        worksheet3.write(gt_idx, c_idx + 1, col_total)
    worksheet3.write(gt_idx, len(date_cols) + 1, pivot_df.loc['Grand Total', 'Grand Total'])

print("Process completed successfully.")
