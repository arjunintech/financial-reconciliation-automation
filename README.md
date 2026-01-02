# Automated Financial Reconciliation System

## 📌 Project Overview
A Python-based automation tool designed to streamline the financial reconciliation process. This system automatically compares **Internal General Ledgers** against **External Bank Statements** to identify discrepancies, missing transactions, and matching records.

## 🚀 The Business Problem
In traditional finance roles, reconciling thousands of transactions manually is:
* **Time-Consuming:** Can take 4+ hours for large ledgers.
* **Error-Prone:** Human eyes often miss small discrepancies.
* **Repetitive:** The logic (match Ref ID -> check Amount) never changes.

## 🛠️ The Solution
This script reduces the process to **< 10 seconds** by:
1.  **Ingesting Data:** Reads `.xlsx` files using Pandas.
2.  **Standardizing:** Cleans and formats diverse column names (e.g., 'Withdrawal/Deposit' vs 'Net Amount').
3.  **Matching Logic:** Performs an outer join on unique Transaction IDs.
4.  **Reporting:** Generates an Excel report with separate sheets for `Matched` and `Action Required` items.

## 💻 Tech Stack
* **Python 3.10+**
* **Pandas:** For high-performance data manipulation.
* **OpenPyXL:** For Excel file handling and formatting.

## 📂 File Structure
* `reconcile.py`: The main engine containing the matching logic.
* `generate_data.py`: A utility script to generate dummy financial data for testing.
* `company_ledger.xlsx`: Sample internal data.
* `bank_statement.xlsx`: Sample bank data (with intentional discrepancies).

## 🔧 How to Run
1.  Install dependencies:
    ```bash
    pip install pandas openpyxl
    ```
2.  Generate fresh dummy data (optional):
    ```bash
    python generate_data.py
    ```
3.  Run the reconciliation:
    ```bash
    python reconcile.py
    ```
4.  Check the output file: `Reconciliation_Report_FINAL.xlsx`
