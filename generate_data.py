import pandas as pd
import random

# --- CONFIGURATION ---
NUM_TRANSACTIONS = 50  # We will generate 50 rows of data

def generate_dummy_data():
    print(f"Generating {NUM_TRANSACTIONS} fake transactions...")
    
    # 1. Create Base Data (Perfect Matches)
    data = []
    for i in range(1, NUM_TRANSACTIONS + 1):
        ref_id = f"TXN-{1000 + i}"
        amount = random.randint(100, 5000)
        # 80% chance of being a Deposit (Income), 20% Withdrawal (Expense)
        txn_type = 'Deposit' if random.random() > 0.2 else 'Withdrawal'
        
        data.append({
            'Ref_ID': ref_id,
            'Amount': amount,
            'Type': txn_type,
            'Description': f"Payment for Invoice {i}",
            'Date': '2025-10-15'
        })
    
    # Convert to DataFrame
    df_common = pd.DataFrame(data)

    # --- PREPARE LEDGER (Internal Record) ---
    df_ledger = df_common.copy()
    df_ledger = df_ledger.rename(columns={'Amount': 'Amount', 'Type': 'Type'})
    
    # --- PREPARE BANK STATEMENT (External Record) ---
    df_bank = df_common.copy()
    
    # --- INTRODUCE ERRORS (So the project actually has work to do) ---
    
    # Error 1: Amount Mismatch (The bank has a different amount than the ledger)
    # We change the amount for the first 3 transactions in the Bank file
    df_bank.loc[0:2, 'Amount'] = df_bank.loc[0:2, 'Amount'] + 50 
    
    # Error 2: Missing in Bank (Check was issued but not cashed yet)
    # We delete the last 5 rows from the Bank file
    df_bank = df_bank.iloc[:-5]
    
    # Error 3: Missing in Ledger (Bank fee or interest appearing only in Bank)
    # We add a new row to the Bank file only
    new_row = pd.DataFrame([{
        'Ref_ID': 'BANK-FEE-001', 
        'Amount': 25, 
        'Type': 'Withdrawal', 
        'Description': 'Monthly Service Charge', 
        'Date': '2025-10-30'
    }])
    df_bank = pd.concat([df_bank, new_row], ignore_index=True)

    # --- FORMATTING FOR BANK (Separate Deposit/Withdrawal columns) ---
    # Banks usually split these into two columns, while Ledgers keep one.
    df_bank['Deposit'] = df_bank.apply(lambda x: x['Amount'] if x['Type'] == 'Deposit' else 0, axis=1)
    df_bank['Withdrawal'] = df_bank.apply(lambda x: x['Amount'] if x['Type'] == 'Withdrawal' else 0, axis=1)
    
    # Drop the original 'Amount' and 'Type' from Bank to look realistic
    df_bank = df_bank.drop(columns=['Amount', 'Type'])
    
    # --- SAVE FILES ---
    df_ledger.to_excel('company_ledger.xlsx', index=False)
    df_bank.to_excel('bank_statement.xlsx', index=False)
    
    print("Success! Created 'company_ledger.xlsx' and 'bank_statement.xlsx'")

if __name__ == "__main__":
    generate_dummy_data()