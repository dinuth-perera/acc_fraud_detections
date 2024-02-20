import pandas as pd

# Sample data for invoice data
invoice_data = {
    "SequenceNumber": [1, 2, 3, 5, 6, 7, 8],
    "InvoiceNumber": ["INV001", "INV002", "INV003", "INV005", "INV006", "INV007", "INV008"],
}

# Create a DataFrame from the sample data
invoice_df = pd.DataFrame(invoice_data)

# Save the DataFrame to an Excel file
invoice_df.to_excel("invoice_data.xlsx", index=False)
