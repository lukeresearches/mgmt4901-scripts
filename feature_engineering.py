import pandas as pd
import os

data_dir = "/Users/decosteluke/Dropbox/ACademic  Teaching - Dalhousie/2025-05 - MGMT 4901 Async/Data Files"
merged_path = os.path.join(data_dir, "ALL_MERGED.xlsx")
features_path = os.path.join(data_dir, "TEAM_FEATURES.xlsx")

df = pd.read_excel(merged_path)

# List of BComm/BMgmt/derivative degrees (case-insensitive match, ignoring trailing *)
bcomm_bmgmt_degrees = [
    "bachelor of management",
    "bachelor of commerce co-op",
    "bachelor of management*",
    "bachelor of commerce co-op*"
]

def is_bcomm_or_bmgmt(degree):
    if pd.isna(degree):
        return 0
    degree_clean = degree.lower().replace('*', '').strip()
    for deg in bcomm_bmgmt_degrees:
        if degree_clean == deg.replace('*', '').strip():
            return 1
    return 0

df["is_bcomm_or_bmgmt"] = df["Degree"].apply(is_bcomm_or_bmgmt)

df["is_entrepreneurship_major"] = df["Major"].str.lower().str.contains("entrepreneurship", na=False).astype(int)

# Save the feature-engineered file
df.to_excel(features_path, index=False)
print(f"Feature-engineered file saved as {features_path}")
