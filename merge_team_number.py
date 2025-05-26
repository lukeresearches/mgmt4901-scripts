import pandas as pd

# === CONFIGURATION ===
# Update these filenames if needed
QUIZ_FILE = "2 Capstone Team Infrastructure Establishment (Group) - Attempt Details_CLEANED.xlsx"  # cleaned quiz file
CLASS_LIST_FILE = "00 Class List.xlsx"
OUTPUT_FILE = "cleaned_quiz_with_teams.xlsx"

# Read the cleaned quiz file
quiz_df = pd.read_excel(QUIZ_FILE)

# Read the class list file
class_list_df = pd.read_excel(CLASS_LIST_FILE)

# Merge on 'username', keeping only quiz rows
df_merged = pd.merge(quiz_df, class_list_df[['username', 'Team Number']], on='username', how='left')

# Save to a new Excel file
df_merged.to_excel(OUTPUT_FILE, index=False)

print(f"Merged file saved as {OUTPUT_FILE}")
