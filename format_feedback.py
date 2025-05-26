import pandas as pd
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def format_feedback():
    try:
        # Load evaluation output
        logging.info("Loading evaluation output...")
        eval_df = pd.read_csv("MGMT4901_3A_Evaluation_Output.csv")
        
        # Load class list
        logging.info("Loading class list...")
        class_df = pd.read_excel("00 Class List.xlsx")
        
        # Convert to wide format
        logging.info("Converting to wide format...")
        # First create separate DataFrames for feedback and scores
        feedback_df = eval_df[['username', 'rubric_category', 'feedback']].copy()
        scores_df = eval_df[['username', 'rubric_category', 'score']].copy()
        
        # Pivot both DataFrames
        feedback_wide = feedback_df.pivot(index='username', columns='rubric_category', values='feedback')
        scores_wide = scores_df.pivot(index='username', columns='rubric_category', values='score')
        
        # Add suffixes to column names
        feedback_wide.columns = [f"{col}_feedback" for col in feedback_wide.columns]
        scores_wide.columns = [f"{col}_score" for col in scores_wide.columns]
        
        # Combine feedback and scores
        wide_df = pd.concat([feedback_wide, scores_wide], axis=1)
        
        # Merge with class list
        logging.info("Merging with class list...")
        # Keep only the columns we need from class list
        class_meta = class_df[['username', 'E-Mail Address', 'First Name', 'Team Number']].copy()
        
        # Merge on username
        final_df = pd.merge(class_meta, wide_df, on='username', how='left')
        
        # Sort columns to put metadata at the start
        meta_cols = ['username', 'E-Mail Address', 'First Name', 'Team Number']
        feedback_score_cols = [col for col in final_df.columns if col not in meta_cols]
        final_df = final_df[meta_cols + feedback_score_cols]
        
        # Save to CSV
        logging.info("Saving final formatted feedback...")
        final_df.to_csv("3A Final_Formatted_Feedback.csv", index=False)
        logging.info("Script completed successfully!")
        
    except Exception as e:
        logging.error(f"Error during processing: {str(e)}")
        raise

if __name__ == "__main__":
    format_feedback()
