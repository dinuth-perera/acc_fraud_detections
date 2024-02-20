import pandas as pd

def detect_sequence_mismatch(data):
    """
    Detects sequence number mismatches in the DataFrame.

    Args:
    - data (DataFrame): Input DataFrame containing the invoice data.

    Returns:
    - List of tuples containing the index and invoice number of mismatched sequences.
    """
    mismatches = []
    prev_sequence = None

    for index, row in data.iterrows():
        current_sequence = row['SequenceNumber']
        if prev_sequence is not None and current_sequence != prev_sequence + 1:
            mismatches.append((index, row['InvoiceNumber']))
        prev_sequence = current_sequence

    return mismatches

# Replace 'invoice_data.xlsx' with the actual file name
invoice_data = pd.read_excel('invoice_data.xlsx')

# Assuming 'SequenceNumber' and 'InvoiceNumber' are the column names
# Modify these column names according to your Excel file
mismatched_invoices = detect_sequence_mismatch(invoice_data)

if len(mismatched_invoices) > 0:
    print("Sequence number mismatches found:")
    for index, invoice_number in mismatched_invoices:
        print(f"Index: {index}, Invoice Number: {invoice_number}")
    
    # Create a DataFrame for the mismatches
    mismatch_df = pd.DataFrame(mismatched_invoices, columns=['Index', 'InvoiceNumber'])
    
    # Save the mismatches to an Excel file
    mismatch_df.to_excel('sequence_number_mismatches.xlsx', index=False)
    print("Mismatched invoices saved to 'sequence_number_mismatches.xlsx'.")
else:
    print("No sequence number mismatches found.")
