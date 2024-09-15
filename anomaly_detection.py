import pandas as pd
import numpy as np
from scipy.stats import zscore

def detect_anomalies(data, threshold=3):
    """
    Detect anomalies using Z-score method.

    Parameters:
    - data: List or array of transaction amounts
    - threshold: Z-score value above which data points are considered anomalies

    Returns:
    - A DataFrame with original data and anomaly flag
    """
    # Convert data to numpy array
    data_array = np.array(data).reshape(-1, 1)
    
    # Calculate Z-scores
    z_scores = zscore(data_array)
    
    # Identify anomalies
    anomalies = np.abs(z_scores) > threshold
    
    return anomalies.flatten()

def main(excel_file, sheet_name, amount_column, threshold=3):
    # Load Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Extract the column with transaction amounts
    amounts = df[amount_column]
    
    # Detect anomalies
    anomalies = detect_anomalies(amounts, threshold)
    
    # Add anomaly flag to the DataFrame
    df['Anomaly'] = anomalies
    
    # Save results to a new Excel file
    output_file = 'anomaly_detection_results.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"Anomaly detection results saved to {output_file}")

# Example usage
if __name__ == "__main__":
    excel_file = 'your_file.xlsx'         # Excel file path
    sheet_name = 'Sheet1'                # sheet name
    amount_column = 'TransactionAmount'  #column name for transaction amounts
    threshold = 3                        # Z-score threshold for anomaly detection
    main(excel_file, sheet_name, amount_column, threshold)
