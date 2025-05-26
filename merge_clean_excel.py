import os
import pandas as pd

# Directory containing Excel files
data_dir = os.path.dirname(os.path.abspath(__file__))

# List all Excel files in the directory
excel_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx') or f.endswith('.xls')]

# Read and clean each file (add your cleaning logic here)
dataframes = []
for file in excel_files:
    df = pd.read_excel(os.path.join(data_dir, file))
    # --- PLACEHOLDER: Add cleaning steps here ---
    dataframes.append(df)

# Merge all DataFrames (simple vertical concat)
if dataframes:
    merged = pd.concat(dataframes, ignore_index=True)
    merged.to_excel(os.path.join(data_dir, 'merged_output.xlsx'), index=False)
    print(f"Merged {len(dataframes)} files into merged_output.xlsx")
else:
    print("No Excel files found in the directory.")
