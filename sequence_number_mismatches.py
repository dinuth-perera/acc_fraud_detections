import pandas as pd

def detect_sequence_mismatch(data):
    """
    Detects sequence number mismatches in the DataFrame.

    Args:
    - data (DataFrame): Input DataFrame containing the invoice data.

    Returns:
    - List of tuples containing the index and invoice number of mismatched sequences.
    """
    return [(index, row['InvoiceNumber']) for index, row in data.iterrows()
            if index > 0 and row['SequenceNumber'] != data.at[index - 1, 'SequenceNumber'] + 1]

# Replace 'invoice_data.xlsx' with the actual file name
invoice_data = pd.read_excel('invoice_data.xlsx')

# Detect mismatched sequences
mismatched_invoices = detect_sequence_mismatch(invoice_data)

if mismatched_invoices:
    print("Sequence number mismatches found:")
    for index, invoice_number in mismatched_invoices:
        print(f"Index: {index}, Invoice Number: {invoice_number}")
    
    # Create a DataFrame for the mismatches and save to Excel
    pd.DataFrame(mismatched_invoices, columns=['Index', 'InvoiceNumber']).to_excel(
        'sequence_number_mismatches.xlsx', index=False
    )
    print("Mismatched invoices saved to 'sequence_number_mismatches.xlsx'.")
else:
    print("No sequence number mismatches found.")
