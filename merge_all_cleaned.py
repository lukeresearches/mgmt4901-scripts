import pandas as pd
import glob
import os
from functools import reduce

data_dir = os.path.dirname(os.path.abspath(__file__))

# Read class list, normalize username column to lower
class_list = pd.read_excel(os.path.join(data_dir, "00 Class List.xlsx"))
# Find the username column (case-insensitive)
class_username_col = [col for col in class_list.columns if col.lower() == "username"][0]
class_list["username"] = class_list[class_username_col].astype(str).str.lower()

# Find the name column (first column that is not 'username')
name_col = [col for col in class_list.columns if col != class_username_col][0]
# Check for duplicate names
duplicates = class_list[class_list.duplicated(subset=[name_col], keep=False)]
if not duplicates.empty:
    print(f"Warning: Duplicate names found in class list (column '{name_col}'):")
    print(duplicates[[name_col, 'username']])

# Read all cleaned quiz files
quiz_files = glob.glob(os.path.join(data_dir, "*_CLEANED.xlsx"))
quiz_dfs = []
for f in quiz_files:
    df = pd.read_excel(f)
    df["username"] = df["username"].astype(str).str.lower()
    quiz_dfs.append(df)

# Merge all quiz dataframes on username (wide format)
if quiz_dfs:
    quiz_merged = reduce(lambda left, right: pd.merge(left, right, on="username", how="outer"), quiz_dfs)  # Merge all quiz dataframes on username (wide format)
else:
    quiz_merged = pd.DataFrame(columns=["username"])

# Merge with class list
final_merged = pd.merge(class_list, quiz_merged, on="username", how="outer")

# Save the merged file
final_merged.to_excel(os.path.join(data_dir, "ALL_MERGED.xlsx"), index=False)
print("Merged file saved as ALL_MERGED.xlsx")
