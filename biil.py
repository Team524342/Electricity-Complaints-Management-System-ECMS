import pandas as pd

# Load the Excel file
file_path = "data/Electricity_Bills_3Months.xlsx"  # Make sure the file is in the same directory
df = pd.read_excel(file_path)

def check_payment_status(customer_id):
    # Filter records for the given customer
    customer_data = df[df["Customer ID"] == customer_id]
    
    if customer_data.empty:
        return "no_records"  # No records found for the customer
    
    # Check if all months are paid
    for index, row in customer_data.iterrows():
        if row['Payment Status'].strip().lower() != 'paid':
            return "unpaid"  # Found an unpaid bill
    
    return "paid"  # All bills are paid

# Example usage
#status = check_payment_status("CUST006")  # Replace with the desired Customer ID
#print(status)