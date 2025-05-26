import os
import pandas as pd

def clean_quiz_file(filepath):
    df = pd.read_excel(filepath)
    # Make all columns lower-case and strip spaces for consistency
    df.columns = [col.lower().strip() for col in df.columns]
    # 1. Drop explainer rows (where section # is not empty)
    if 'section #' in df.columns:
        df = df[df['section #'].isna()]
    # 2. Keep only relevant columns
    keep_cols = ['username', 'q #', 'q type', 'q text', 'answer', 'answer match']
    df = df[[col for col in keep_cols if col in df.columns]]

    # 3. Pivot/widen data by Q Type using a list of dicts
    user_data_list = []
    usernames = df['username'].unique()
    for user in usernames:
        user_rows = df[df['username'] == user]
        user_data = {'username': user}
        for qnum in user_rows['q #'].dropna().unique():
            qrows = user_rows[user_rows['q #'] == qnum]
            qtype = qrows['q type'].iloc[0]
            qtext = qrows['q text'].iloc[0]
            if qtype == 'WR':
                # Written Response
                user_data[qtext] = qrows['answer'].iloc[0]
            elif qtype == 'MC':
                # Multiple Choice - get checked answer
                checked = qrows[qrows['answer match'].str.lower() == 'checked']
                if not checked.empty:
                    user_data[qtext] = checked['answer'].iloc[0]
            elif qtype == 'M-S':
                # Multi-Select - make a column for each option
                for _, row in qrows.iterrows():
                    colname = f"{qtext}: {row['answer']}"
                    user_data[colname] = 1 if str(row['answer match']).lower() == 'checked' else 0
            elif qtype == 'MSA':
                # Each MSA row gets its own column, value from 'answer match'
                for idx, (_, msa_row) in enumerate(qrows.iterrows()):
                    msa_col = f"{qtext} {idx+1}" if len(qrows) > 1 else qtext
                    user_data[msa_col] = msa_row['answer match']
        user_data_list.append(user_data)
    wide = pd.DataFrame(user_data_list)
    return wide

if __name__ == "__main__":
    data_dir = os.path.dirname(os.path.abspath(__file__))
    # Example: Clean all quiz files in the directory except the class list
    for fname in os.listdir(data_dir):
        if fname.endswith('.xlsx') and not fname.startswith('00 Class List'):
            print(f"Cleaning {fname}...")
            cleaned = clean_quiz_file(os.path.join(data_dir, fname))
            outname = fname.replace('.xlsx', '_CLEANED.xlsx')
            cleaned.to_excel(os.path.join(data_dir, outname), index=False)
            print(f"Saved cleaned file: {outname}")
