import pandas as pd
import numpy as np

# Create routes database
routes = pd.DataFrame({
    'ROUTE_CODE': ['LJ01', 'LJ02', 'MB01', 'KP01', 'CE01'],
    'ZIP_CODE': ['1000', '1001', '2000', '6000', '3000'],
    'CITY': ['Ljubljana center', 'Ljubljana Šiška', 'Maribor', 'Koper', 'Celje']
})
routes.to_excel('input/routes_database.xlsx', index=False)
print("Created routes database")

# Create streets database (optional)
streets = pd.DataFrame({
    'ROUTE_CODE': ['LJ01', 'LJ01', 'LJ02', 'LJ02', 'MB01'],
    'STREET_NAME': ['Slovenska cesta', 'Čopova ulica', 'Celovška cesta', 'Šišenska cesta', 'Glavni trg']
})
streets.to_excel('input/street_database.xlsx', index=False)
print("Created streets database")

# Create sample manifest
np.random.seed(42)  # For reproducibility

# Generate 50 sample shipments
manifest = pd.DataFrame({
    'HWB': [f'SHP{i:06d}' for i in range(1, 51)],
    'CONSIGNEE_NAME': [f'Customer {i}' for i in range(1, 51)],
    'CONSIGNEE_ZIP': np.random.choice(['1000', '1001', '2000', '3000', '6000'], 50),
    'CONSIGNEE_STREET': [f'Address {i}' for i in range(1, 51)],
    'WEIGHT': np.random.uniform(1, 100, 50).round(1),
    'VOLUMETRIC_WEIGHT': np.random.uniform(1, 200, 50).round(1),
    'PIECES': np.random.randint(1, 10, 50)
})
manifest.to_excel('input/daily_manifest.xlsx', index=False)
print("Created sample manifest")
