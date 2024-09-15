import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from collections import Counter

def get_leading_digit(number):
    """Extract the leading digit from a number."""
    while number >= 10:
        number //= 10
    return number

def benford_distribution():
    """Return the expected Benford's Law distribution for leading digits."""
    return {i: np.log10(1 + 1 / i) for i in range(1, 10)}

def analyze_benford(data):
    """Analyze data to compare with Benford's Law."""
    # Extract leading digits from the data
    digits = [get_leading_digit(int(str(x).strip())) for x in data if str(x).strip().isdigit()]
    
    # Count occurrences of each leading digit
    counts = Counter(digits)
    
    # Calculate frequencies
    total = sum(counts.values())
    observed_distribution = {i: counts[i] / total for i in range(1, 10)}
    
    # Get Benford's Law distribution
    expected_distribution = benford_distribution()
    
    return observed_distribution, expected_distribution

def plot_distributions(observed, expected):
    """Plot the observed vs expected Benford's Law distributions."""
    digits = range(1, 10)
    observed_values = [observed.get(d, 0) for d in digits]
    expected_values = [expected.get(d, 0) for d in digits]

    plt.figure(figsize=(10, 6))
    plt.bar(digits, observed_values, width=0.4, align='center', label='Observed', alpha=0.6)
    plt.plot(digits, expected_values, marker='o', color='r', label='Expected (Benford\'s Law)', linestyle='--')
    
    plt.xlabel('Leading Digit')
    plt.ylabel('Frequency')
    plt.title('Benford\'s Law Analysis')
    plt.xticks(digits)
    plt.legend()
    plt.grid(True)
    plt.show()

def main(excel_file, sheet_name):
    # Load Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Flatten the DataFrame to a single series
    data = df.values.flatten()
    
    # Analyze Benford's Law
    observed, expected = analyze_benford(data)
    
    # Plot the distributions
    plot_distributions(observed, expected)

# Example usage
if __name__ == "__main__":
    excel_file = 'data.xlsx'  # Excel file path
    sheet_name = 'Sheet1'          # sheet name
    main(excel_file, sheet_name)
