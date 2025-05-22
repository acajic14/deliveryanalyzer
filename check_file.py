import pandas as pd
import os

# Print current working directory to verify where we're looking
print(f"Current working directory: {os.getcwd()}")

# Check if the file exists
file_path = 'input/routes_database.xlsx'
print(f"File exists: {os.path.exists(file_path)}")

# Try to load the file
try:
    routes_df = pd.read_excel(file_path)
    print(f"Successfully loaded the file with {len(routes_df)} rows")
    print("\nFirst few rows:")
    print(routes_df.head())
    print("\nColumn names:", routes_df.columns.tolist())
except Exception as e:
    print(f"Error loading file: {str(e)}")
