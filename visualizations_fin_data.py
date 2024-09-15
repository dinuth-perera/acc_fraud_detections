import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import numpy as np

# Load data from Excel file
def load_data(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)

# Generate time series plot
def plot_time_series(df, date_column, value_column):
    plt.figure(figsize=(14, 7))
    plt.plot(df[date_column], df[value_column], marker='o', linestyle='-')
    plt.title(f'Time Series of {value_column}')
    plt.xlabel('Date')
    plt.ylabel(value_column)
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# Generate correlation heatmap
def plot_correlation_heatmap(df):
    plt.figure(figsize=(12, 8))
    correlation_matrix = df.corr()
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', vmin=-1, vmax=1)
    plt.title('Correlation Heatmap')
    plt.show()

# Generate scatter plot with trend line
def plot_scatter_with_trend(df, x_column, y_column):
    plt.figure(figsize=(10, 6))
    sns.regplot(x=x_column, y=y_column, data=df, scatter_kws={'s':50}, line_kws={'color':'red'})
    plt.title(f'Scatter Plot with Trend Line: {x_column} vs {y_column}')
    plt.xlabel(x_column)
    plt.ylabel(y_column)
    plt.grid(True)
    plt.show()

# Identify and plot anomalies
def plot_anomalies(df, date_column, value_column, threshold=1.5):
    df['Moving_Avg'] = df[value_column].rolling(window=12).mean()
    df['Moving_Std'] = df[value_column].rolling(window=12).std()
    df['Anomaly'] = (df[value_column] > df['Moving_Avg'] + (threshold * df['Moving_Std'])) | \
                    (df[value_column] < df['Moving_Avg'] - (threshold * df['Moving_Std']))

    plt.figure(figsize=(14, 7))
    plt.plot(df[date_column], df[value_column], label='Value', color='blue')
    plt.plot(df[date_column], df['Moving_Avg'], label='Moving Average', color='green')
    plt.scatter(df[df['Anomaly']][date_column], df[df['Anomaly']][value_column], color='red', label='Anomalies', marker='o')
    plt.title('Anomalies Detection')
    plt.xlabel('Date')
    plt.ylabel(value_column)
    plt.legend()
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# Generate interactive 3D scatter plot
def plot_3d_scatter(df, x_column, y_column, z_column):
    fig = px.scatter_3d(df, x=x_column, y=y_column, z=z_column, color=df[z_column], title=f'3D Scatter Plot: {x_column}, {y_column}, {z_column}')
    fig.show()

# Main function
def main():
    file_path = 'financial_data.xlsx'  #file path
    sheet_name = 'Sheet1'  #sheet name
    df = load_data(file_path, sheet_name)
    
    # Assume the DataFrame has columns 'Date', 'Revenue', 'Expense', etc.
    date_column = 'Date'
    value_column = 'Revenue'  # Change this as needed

    # Generate visualizations
    plot_time_series(df, date_column, value_column)
    plot_correlation_heatmap(df)
    plot_scatter_with_trend(df, 'Revenue', 'Expense')  # Adjust columns as needed
    plot_anomalies(df, date_column, value_column)
    plot_3d_scatter(df, 'Revenue', 'Expense', 'Profit')  # Adjust columns as needed

if __name__ == "__main__":
    main()
