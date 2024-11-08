import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Access the Monzo API credentials
MONZO_ACCESS_TOKEN = os.getenv("MONZO_ACCESS_TOKEN")
MONZO_ACCOUNT_ID = os.getenv("MONZO_ACCOUNT_ID")


# Set up the API URL and headers
MONZO_API_URL = "https://api.monzo.com/transactions"
headers = {
    "Authorization": f"Bearer {MONZO_ACCESS_TOKEN}"
}

# Step 1: Test authentication
def test_authentication():
    url = "https://api.monzo.com/ping/whoami"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        print("Authentication successful.")
    else:
        print(f"Authentication failed with status code: {response.status_code}, response: {response.json()}")

# Step 2: Fetch all transactions for the specified account
def fetch_transactions(account_id):
    url = MONZO_API_URL
    params = {
        "account_id": account_id,
    }
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()
        return data.get("transactions", [])
    else:
        print(f"Error fetching transactions: {response.status_code}")
        return []

# Step 3: Process and categorize transactions for the current month
def categorize_transactions(transactions):
    categories = {}
    all_categories = ['eating_out', 'groceries', 'bills', 'shopping', 'entertainment', 'transport', 'uncategorized']  # List all categories you want to track

    # Initialize each category with an empty list
    for category in all_categories:
        formatted_category = category.replace('_', ' ').title()  # Capitalize each word and replace underscores with spaces
        categories[formatted_category] = []

    # Process each transaction
    for transaction in transactions:
        amount = transaction['amount'] / 100.0  # Convert pence to pounds
        category = transaction.get('category', 'uncategorized').replace('_', ' ').title()  # Format category name

        # Only consider expenses (negative amounts)
        if amount < 0:
            categories[category].append({
                'description': transaction.get('description', 'No description')[:18],
                'amount': -amount  # Negative amount to represent expenditure
            })
    
    return categories

# Step 4: Save results to Excel with categories as columns and descriptions + amounts under each
def save_to_excel(categories):
    # Prepare the list of categories
    category_names = list(categories.keys())
    all_data = []
    total_expenditure = 0  # Initialize total expenditure

    # Iterate through each category
    for category in category_names:
        descriptions = []
        amounts = []
        total_for_category = 0  # Initialize category total

        for transaction in categories[category]:
            descriptions.append(transaction['description'])
            amounts.append(transaction['amount'])
            total_for_category += transaction['amount']

        # Append a description for the total row and the total amount for the category
        descriptions.append('∑ ' + category)
        amounts.append(f'£{total_for_category:.2f}')
        
        # Add the category total to the overall total expenditure
        total_expenditure += total_for_category

        # Add category data to the all_data list
        all_data.append({
            'Description': descriptions,
            'Amount': amounts
        })

    # Prepare columns for the DataFrame
    all_columns = []
    for category in category_names:
        all_columns.append(f'{category}')  # Category Description column
        all_columns.append(f'£ {category}')  # Category Amount column

    # Initialize DataFrame
    df = pd.DataFrame(columns=all_columns)

    # Find the maximum number of rows
    max_rows = max([len(category_data['Description']) for category_data in all_data])

    # Populate DataFrame with each category's description and amount
    for i in range(max_rows):  # Iterate through rows
        row = []
        for category_data in all_data:
            row.append(category_data['Description'][i] if i < len(category_data['Description']) else '')
            row.append(category_data['Amount'][i] if i < len(category_data['Amount']) else '')
        # Ensure the row has the same number of elements as there are columns
        while len(row) < len(all_columns):
            row.append('')

        # Add the row to the DataFrame
        df.loc[i] = row

    # Add the total expenditure row at the bottom (only once)
    total_row = ['Total Expenditure', f'£{total_expenditure:.2f}'] + [''] * (len(all_columns) - 2)
    df.loc[max_rows] = total_row

    # Save to Excel file
    excel_filename = 'monzo_transactions.xlsx'
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Transactions', index=False)

    # Get the workbook and sheet to modify column widths
    workbook = writer.book
    sheet = workbook['Transactions']

    # Set column width for each column to 140px (approximately 10 characters wide)
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name (e.g., 'A', 'B', etc.)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # Adding extra space for padding
        sheet.column_dimensions[column].width = 140 / 7  # Adjust based on your requirement, 140px converted to column width

    # Save the workbook with adjusted column widths
    workbook.save(excel_filename)

    print("Results saved to 'monzo_transactions.xlsx'.")

    # Open the file automatically (works on most OS)
    open_file(excel_filename)
# Function to open the saved Excel file
def open_file(filepath):
    try:
        # This should work on Windows, macOS, and Linux
        if os.name == 'posix':  # macOS or Linux
            os.system(f'open {filepath}')
        elif os.name == 'nt':  # Windows
            os.startfile(filepath)
        else:
            print("Unable to open the file. Please open it manually.")
    except Exception as e:
        print(f"Error opening the file: {e}")
        


# Step 5: Main function to fetch, categorize, and calculate total expenditure
def main():
    # First, test the authentication
    test_authentication()
    
    # Fetch transactions for the current month
    transactions = fetch_transactions(MONZO_ACCOUNT_ID)
    
    if transactions:
        categories = categorize_transactions(transactions)
        
        # Display categorized expenditures
        print("\nTotal Expenditure by Category (Current Month):")
        for category, transactions in categories.items():
            category_total = sum([t['amount'] for t in transactions])  # Sum the amounts for each category
            print(f"{category.capitalize()}: £{category_total:.2f}")
        
        # Save the results to an Excel file
        save_to_excel(categories)
    else:
        print("No transactions found.")

# Run the main function
if __name__ == "__main__":
    main()