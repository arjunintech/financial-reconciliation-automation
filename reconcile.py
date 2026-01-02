import pandas as pd
import datetime

# --- CONFIGURATION ---
LEDGER_FILE = 'company_ledger.xlsx'
BANK_FILE = 'bank_statement.xlsx'
OUTPUT_FILE = 'Reconciliation_Report_FINAL.xlsx'

def load_data():
    print("Loading files...")
    # Load Company Ledger
    df_ledger = pd.read_excel(LEDGER_FILE)
    
    # Load Bank Statement
    df_bank = pd.read_excel(BANK_FILE)
    
    # Standardize Column Names for easier matching
    # We rename columns to match a standard format
    df_ledger = df_ledger.rename(columns={'Reference_ID': 'Ref_ID', 'Amount': 'Ledger_Amount'})
    
    # Clean Data: Ensure Ref_ID is a string (remove decimals if any)
    df_ledger['Ref_ID'] = df_ledger['Ref_ID'].astype(str).str.strip()
    df_bank['Ref_ID'] = df_bank['Ref_ID'].astype(str).str.strip()
    
    return df_ledger, df_bank

def reconcile_data(df_ledger, df_bank):
    print("Running reconciliation logic...")
    
    # MERGE the two files based on the Unique Reference ID
    # 'outer' join keeps ALL records from both files so we can see what is missing
    merged_df = pd.merge(df_ledger, df_bank, on='Ref_ID', how='outer', indicator=True)
    
    # LOGIC 1: Check if Amount Matches
    # (Assuming 'Ledger_Amount' should match 'Deposit' for income or 'Withdrawal' for expense)
    # For simplicity, let's assume we are comparing a net amount column. 
    # If your bank has separate Deposit/Withdrawal, create a Net Amount column first.
    
    # Let's handle a simple case: Calculate difference
    # Fills NaN with 0 to avoid errors
    merged_df['Ledger_Amount'] = merged_df['Ledger_Amount'].fillna(0)
    
    # Create a 'Bank_Amount' column (Combine Deposit and Withdrawal for comparison)
    # Note: Adjust this logic based on your actual Excel structure
    merged_df['Bank_Amount'] = merged_df['Deposit'].fillna(0) - merged_df['Withdrawal'].fillna(0)
    
    # Calculate Difference
    merged_df['Difference'] = merged_df['Ledger_Amount'] - merged_df['Bank_Amount']
    
    # LOGIC 2: Status Flagging
    def get_status(row):
        if row['_merge'] == 'left_only':
            return "Missing in Bank"
        elif row['_merge'] == 'right_only':
            return "Missing in Ledger"
        elif row['Difference'] != 0:
            return "Amount Mismatch"
        else:
            return "Matched"

    merged_df['Reconciliation_Status'] = merged_df.apply(get_status, axis=1)
    
    return merged_df

def save_report(merged_df):
    print(f"Saving report to {OUTPUT_FILE}...")
    
    # Separate the data into sheets for the accountant
    matched_records = merged_df[merged_df['Reconciliation_Status'] == 'Matched']
    discrepancies = merged_df[merged_df['Reconciliation_Status'] != 'Matched']
    
    # Write to Excel with multiple sheets
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        discrepancies.to_excel(writer, sheet_name='Action Required', index=False)
        matched_records.to_excel(writer, sheet_name='Successfully Matched', index=False)
        
    print("Success! Report generated.")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    try:
        ledger, bank = load_data()
        result = reconcile_data(ledger, bank)
        save_report(result)
    except Exception as e:
        print(f"An error occurred: {e}")