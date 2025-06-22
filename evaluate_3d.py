import pandas as pd
import openai
from dotenv import load_dotenv
import os
import json
import logging
from typing import Dict, List, Tuple

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuration will be loaded by load_config() function

# Configuration
CONFIG = {
    "submission_file": "/Users/decosteluke/Dropbox/ACademic  Teaching - Dalhousie/2025-05 - MGMT 4901 Async/Data Files/00 Process/3D – Prototype Testing  Quiz Process.xlsx",
    "rubric_file": "/Users/decosteluke/Dropbox/ACademic  Teaching - Dalhousie/2025-05 - MGMT 4901 Async/Data Files/00 Process/2025_05_Rubric_Table.xlsx",
    "output_file": "MGMT4901_3D_Evaluation_OutputR1.csv",
    "id_column": "username",
    "rubric_sheet": "Rubric",
    "model": "gpt-4"
}

# System and prompt templates - Token efficient version
SYSTEM_PROMPT = """You are a helpful teaching assistant evaluating student submissions in an entrepreneurship course focused on prototype testing. 

Your evaluation should follow these principles:
- Be constructive, warm, plainspoken, and grounded
- Keep all feedback extremely brief - EXACTLY ONE SENTENCE (maximum 25 words)
- ALWAYS make explicit references to the student's actual submission by directly quoting specific words or phrases they used
- Mention one strength and one area for improvement
- Score fairly (15-16 for adequate work, 17-18 for strong work, only deduct for real issues)

For Assignment Context - Prototype Testing:
Students were expected to build a tangible prototype (e.g., mock-up, landing page, AI demo) that made key assumptions testable, and then conduct actual tests with users to validate or invalidate their hypotheses from their Theory of Value Canvas. Prototypes should prioritize learning over polish, following Minimum Viable Product principles.

Use these style guidelines:
- Extremely concise (ONE SENTENCE ONLY - MAXIMUM 25 WORDS)
- Include at least one direct quote or specific reference from the student's work in every piece of feedback
- Be specific and plainspoken, avoid generic praise
- Use active voice and curious phrasing

IMPORTANT: Always return a valid JSON object with no additional text or explanations outside the JSON."""

RUBRIC_PROMPT_TEMPLATE = """Evaluate the student's prototype testing submission for the {category} category.

**Rubric Items:**
{rubric_items}

**Professor's Guidance:**
{professor_guidance}

**Student's Response:**
{responses}

Respond with JSON only:
{{"rubric_category": "{category}", "feedback": "Your detailed feedback here", "score": "XX"}}
"""

# Summary prompt template
SUMMARY_PROMPT_TEMPLATE = """Provide an overall summary evaluation of the student's prototype testing submission, focusing specifically on their prototype development and testing approach.

**Professor's Guidance:**
{professor_guidance}

**Student's Submissions:**
{responses}

**Prototype Testing Information:**
{prototype_testing_content}

Your summary feedback must explicitly mention:
1. The type of prototype the student created
2. Their testing methodology
3. How they validated/invalidated their hypothesis

Make specific references to what the student wrote by directly quoting their own words.

Respond with JSON only:
{{"feedback": "Your detailed summary feedback here", "score": "XX"}}
"""

def load_rubric() -> Dict[str, List[Dict]]:
    """Load rubric from Excel file and enhance with prototype testing criteria."""
    try:
        # Load rubric data from Excel file
        rubric_df = pd.read_excel(CONFIG["rubric_file"], sheet_name=CONFIG["rubric_sheet"])
        
        # Group by category and convert to list of dictionaries
        rubric_by_category = {}
        for _, row in rubric_df.iterrows():
            category = row['Rubric Category']
            if category not in rubric_by_category:
                rubric_by_category[category] = []
            rubric_by_category[category].append({
                'item': row['Rubric Item']
            })
        
        # Enhance the Hypothesis Testing category with specific prototype testing criteria
        if "Hypothesis Testing" in rubric_by_category:
            # Add prototype testing specific criteria
            prototype_test_criteria = {
                'item': "The test uses a tangible prototype (e.g., mock-up, landing page, AI demo) that makes key assumptions testable, targets a specific hypothesis from their Theory of Value, gathers real user feedback, and emphasizes learning over polish following MVP principles."
            }
            rubric_by_category["Hypothesis Testing"].append(prototype_test_criteria)
            logging.info("Added prototype testing criteria to the Hypothesis Testing rubric category")
        
        return rubric_by_category
        
    except Exception as e:
        logging.error(f"Error loading rubric: {str(e)}")
        raise

def collect_category_responses(student_responses: Dict, category: str) -> str:
    """Collect all responses for a given rubric category."""
    responses = []
    
    # Extract category prefix (e.g., "Evaluation" from "Evaluation / Decision")
    category_prefix = category.split(" ")[0]
    
    for col, response in student_responses.items():
        # Match if full category name is in column OR if the column starts with the category prefix
        if isinstance(response, str) and (category in col or col.startswith(category_prefix)):
            responses.append(f"{col}: {response}")
            
    return "\n".join(responses)

def create_prompt(category: str, rubric_items: List[Dict], student_responses: Dict) -> str:
    """Create the evaluation prompt for OpenAI API."""
    # Format rubric items as bulleted list
    rubric_items_str = "\n".join([f"- {item['item']}" for item in rubric_items])
    
    # Get professor feedback, or use a placeholder if not available
    professor_feedback = student_responses.get('Professor Feedback', '')
    professor_guidance = professor_feedback if professor_feedback else "None provided yet."
    
    # Create the prompt
    prompt = RUBRIC_PROMPT_TEMPLATE.format(
        category=category,
        rubric_items=rubric_items_str,
        professor_guidance=professor_guidance,
        responses=student_responses.get('responses', '')
    )
    
    return prompt

def evaluate_category(student_id: str, category: str, rubric_items: List[Dict], student_responses: Dict) -> Dict:
    """Evaluate a single rubric category for a student and return a dictionary with feedback and score."""
    try:
        # Create the prompt
        prompt = create_prompt(category, rubric_items, student_responses)
        
        # Call OpenAI API
        response = openai.ChatCompletion.create(
            model=CONFIG["model"],
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        
        # Log token usage
        usage = response['usage']
        logging.info(f"[{student_id} | {category}] Token usage — prompt: {usage['prompt_tokens']}, completion: {usage['completion_tokens']}, total: {usage['total_tokens']}")
        
        # Extract JSON from the response with enhanced error handling
        content = response.choices[0].message.content
        
        # Add additional sanitizing to help with common JSON errors
        try:
            # Log the raw content for debugging
            logging.debug(f"Raw API response for {student_id}, category {category}: {content}")
            
            # Check if we have a properly formatted JSON response (must start with { and end with })
            content = content.strip()
            if not (content.startswith('{') and content.endswith('}')):
                logging.warning(f"Response not valid JSON format: {content[:50]}...")
                # Try to extract a JSON object if it exists within the text
                json_start = content.find('{')
                json_end = content.rfind('}')
                
                if json_start >= 0 and json_end > json_start:
                    content = content[json_start:json_end+1]
                    logging.info(f"Extracted JSON content: {content[:50]}...")
                else:
                    raise ValueError("Could not extract valid JSON from response")
            
            result = json.loads(content)
            
            # Validate the response
            if 'rubric_category' not in result or 'feedback' not in result or 'score' not in result:
                logging.warning(f"Missing required fields in response for {student_id}, category {category}")
                return {
                    "rubric_category": category,
                    "feedback": f"Error processing evaluation for {category}. Please try again.",
                    "score": ""
                }
                
        except json.JSONDecodeError as e:
            logging.error(f"JSON parsing error for {student_id}, category {category}: {str(e)}\nContent: {content[:200]}")
            return {
                "rubric_category": category,
                "feedback": f"Error processing evaluation for {category}. Please try again.",
                "score": ""
            }
            
        return result
    except Exception as e:
        logging.error(f"Error processing evaluation for student {student_id}, category {category}: {str(e)}")
        return {
            "rubric_category": category,
            "feedback": f"Error processing evaluation for {category}. Please try again.",
            "score": ""
        }

def create_summary_prompt(student_responses: Dict) -> str:
    """Create the summary prompt for OpenAI API with professor guidance."""
    # Get professor feedback if available
    professor_feedback = student_responses.get('Professor Feedback', '')
    professor_guidance = professor_feedback if professor_feedback else "None provided yet."
    
    # Extract prototype testing specific content for special emphasis
    prototype_testing_content = ""
    
    # Look for columns related to Hypothesis Testing or including 'prototype' keyword
    for col, val in student_responses.items():
        if isinstance(val, str) and val.strip() and (
            "Hypothesis Testing" in col 
            or "prototype" in col.lower() 
            or "test" in col.lower() 
            or col.startswith("Evaluation")
        ):
            prototype_testing_content += f"\n{col}: {val}"
    
    # If no specific prototype testing content was found, use a general message
    if not prototype_testing_content.strip():
        prototype_testing_content = "No specific prototype testing details extracted. Please review all student responses."
    
    # Format the prompt
    all_responses = "\n\n".join([f"{k}: {v}" for k, v in student_responses.items() if k != 'Professor Feedback'])
    prompt = SUMMARY_PROMPT_TEMPLATE.format(
        professor_guidance=professor_guidance,
        responses=all_responses,
        prototype_testing_content=prototype_testing_content
    )
    
    return prompt

def evaluate_all_categories(student_id: str, rubric_by_category: Dict, submission_df: pd.DataFrame, row_index: int) -> Dict:
    """Evaluate all rubric categories for a single student and return a dictionary of results.
    
    Args:
        student_id: Student identifier
        rubric_by_category: Dictionary of rubric categories and items
        submission_df: DataFrame with student submissions
        row_index: Row index for this student in the DataFrame
    """
    # Extract student's responses
    student_row = submission_df.iloc[row_index].to_dict()
    
    # Collect student responses
    student_responses = {}
    for col, val in student_row.items():
        if pd.notna(val) and isinstance(val, str) and val.strip():
            student_responses[col] = val
    
    results = {
        "feedback_by_category": {},
        "score_by_category": {},
        "summary_feedback": ""
    }
    
    # Evaluate each category
    for category, rubric_items in rubric_by_category.items():
        try:
            # Skip empty categories
            if not rubric_items:
                continue
                
            # Get category-specific responses
            category_responses = collect_category_responses(student_responses, category)
            
            # Include all student responses and category-specific responses
            eval_responses = student_responses.copy()
            eval_responses['responses'] = category_responses
            
            # Evaluate the category
            evaluation = evaluate_category(student_id, category, rubric_items, eval_responses)
            
            # Extract feedback and score
            results["feedback_by_category"][category] = evaluation["feedback"]
            results["score_by_category"][category] = evaluation["score"]
            
        except Exception as e:
            logging.error(f"Error evaluating category {category} for student {student_id}: {str(e)}")
            results["feedback_by_category"][category] = f"Error processing evaluation for {category}. Please try again."
            results["score_by_category"][category] = ""
    
    # Generate summary feedback
    try:
        # Call OpenAI API for summary
        summary_prompt = create_summary_prompt(student_responses)
        summary_response = openai.ChatCompletion.create(
            model=CONFIG["model"],
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": summary_prompt}
            ],
            temperature=0.7
        )
        
        # Log token usage for summary
        summary_usage = summary_response['usage']
        logging.info(f"[{student_id} | SUMMARY] Token usage — prompt: {summary_usage['prompt_tokens']}, completion: {summary_usage['completion_tokens']}, total: {summary_usage['total_tokens']}")
        
        # Extract JSON from the summary response with enhanced error handling
        summary_content = summary_response.choices[0].message.content
        
        try:
            # Log the raw content for debugging
            logging.debug(f"Raw summary API response for {student_id}: {summary_content}")
            
            # Check if we have a properly formatted JSON response
            summary_content = summary_content.strip()
            if not (summary_content.startswith('{') and summary_content.endswith('}')):
                logging.warning(f"Summary response not valid JSON format: {summary_content[:50]}...")
                # Try to extract a JSON object if it exists within the text
                json_start = summary_content.find('{')
                json_end = summary_content.rfind('}')
                
                if json_start >= 0 and json_end > json_start:
                    summary_content = summary_content[json_start:json_end+1]
                    logging.info(f"Extracted JSON from summary content: {summary_content[:50]}...")
                else:
                    raise ValueError("Could not extract valid JSON from summary response")
            
            summary_result = json.loads(summary_content)
            results["summary_feedback"] = summary_result.get("feedback", "Error generating summary. Please try again.")
            
        except (json.JSONDecodeError, ValueError) as e:
            logging.error(f"JSON parsing error for {student_id} summary: {str(e)}\nContent: {summary_content[:200]}")
            results["summary_feedback"] = f"Error generating summary. JSON parsing failed: {str(e)[:50]}"
    except Exception as e:
        logging.error(f"Error generating summary for student {student_id}: {str(e)}")
        results["summary_feedback"] = "Error generating summary. Please try again."
    
    return results

def load_config():
    """Load configuration settings."""
    load_dotenv()
    openai.api_key = os.getenv("OPENAI_API_KEY")
    
    if not openai.api_key:
        raise ValueError("OpenAI API key not found. Please set the OPENAI_API_KEY environment variable.")
    
    logging.info("Configuration loaded successfully")

def main(test_mode=True):
    """Main function to evaluate submissions.
    
    Args:
        test_mode (bool): If True, only process the first student in the spreadsheet
    """
    try:
        # Load configuration settings
        load_config()
        
        # Load the rubric
        logging.info("Loading rubric...")
        rubric = load_rubric()
        logging.info(f"Loaded rubric with {sum(len(items) for items in rubric.values())} items across {len(rubric)} categories")
        
        # Load student submissions
        logging.info(f"Loading submissions from {CONFIG['submission_file']}")
        submission_df = pd.read_excel(CONFIG['submission_file'])
        logging.info(f"Loaded {len(submission_df)} submissions")
        
        required_columns = [
            "Summary Feedback", "Total Score", 
            "Capstone Execution Feedback", "Theory Development Feedback", "Hypothesis Development Feedback", 
            "Hypothesis Testing Feedback", "Evaluation / Decision Feedback",
            "Capstone Execution Score (on 20)", "Theory Development Score (on 20)", "Hypothesis Development Score (on 20)", 
            "Hypothesis Testing Score (on 20)", "Evaluation / Decision Score (on 20)"
        ]
        
        # Add any missing columns to the dataframe
        for col in required_columns:
            if col not in submission_df.columns:
                submission_df[col] = ""
                
        # Create a mapping of category names to column names
        feedback_column_map = {
            "Capstone Execution": "Capstone Execution Feedback",
            "Theory Development": "Theory Development Feedback",
            "Hypothesis Development": "Hypothesis Development Feedback",
            "Hypothesis Testing": "Hypothesis Testing Feedback",
            "Evaluation / Decision": "Evaluation / Decision Feedback"
        }
        
        score_column_map = {
            "Capstone Execution": "Capstone Execution Score (on 20)",
            "Theory Development": "Theory Development Score (on 20)",
            "Hypothesis Development": "Hypothesis Development Score (on 20)",
            "Hypothesis Testing": "Hypothesis Testing Score (on 20)",
            "Evaluation / Decision": "Evaluation / Decision Score (on 20)"
        }
        
        # Check if category names and column names exist in the expected format
        all_columns = submission_df.columns.tolist()
        logging.info(f"Available columns: {all_columns}")
        
        # Normalize column names to remove extra spaces around slashes
        normalized_column_map = {}
        for category, column in feedback_column_map.items():
            # Find best matching column
            normalized_col = column.replace(" / ", "/")
            if normalized_col in all_columns:
                normalized_column_map[category] = normalized_col
            elif column in all_columns:
                normalized_column_map[category] = column
            else:
                logging.warning(f"Could not find column matching '{column}' or '{normalized_col}'")
        
        # Use normalized column names if found
        if normalized_column_map:
            feedback_column_map = normalized_column_map
            logging.info(f"Using normalized column names: {feedback_column_map}")
        
        # Do the same for score columns
        normalized_score_map = {}
        for category, column in score_column_map.items():
            normalized_col = column.replace(" / ", "/")
            if normalized_col in all_columns:
                normalized_score_map[category] = normalized_col
            elif column in all_columns:
                normalized_score_map[category] = column
            else:
                logging.warning(f"Could not find column matching '{column}' or '{normalized_col}'")
        
        if normalized_score_map:
            score_column_map = normalized_score_map
            logging.info(f"Using normalized score column names: {score_column_map}")
        
        # Get list of unique student IDs
        student_ids = submission_df[CONFIG["id_column"]].unique()
        logging.info(f"Found {len(student_ids)} students to evaluate")
        
        if len(student_ids) == 0:
            logging.warning("No students found in the spreadsheet!")
            return

        # Process students based on mode
        if test_mode:
            # Process only the first student in test mode
            test_student = student_ids[0]  # First student ID
            logging.info(f"TEST MODE: Processing only the first student: {test_student}")
            student_ids = [test_student]
        else:
            logging.info(f"FULL MODE: Processing all {len(student_ids)} students")
        
        # Process selected students
        for student_id in student_ids:
            # Find the row index for this student
            row_indices = submission_df.index[submission_df[CONFIG["id_column"]] == student_id].tolist()
            if not row_indices:
                logging.warning(f"Could not find row index for student {student_id}, skipping")
                continue
                
            row_index = row_indices[0]
            
            # Evaluate the student
            logging.info(f"Evaluating student: {student_id}")
            try:
                results = evaluate_all_categories(student_id, rubric, submission_df, row_index)
                
                # Update category feedback and scores
                total_score = 0
                score_count = 0
                
                for category, feedback in results["feedback_by_category"].items():
                    if category in feedback_column_map:
                        submission_df.at[row_index, feedback_column_map[category]] = feedback
                
                for category, score in results["score_by_category"].items():
                    if category in score_column_map:
                        try:
                            score_value = float(score)
                            total_score += score_value
                            score_count += 1
                            submission_df.at[row_index, score_column_map[category]] = score_value
                        except (ValueError, TypeError):
                            logging.warning(f"Could not convert score to number: {score}")
                            submission_df.at[row_index, score_column_map[category]] = score
                
                # Calculate and update the total score (sum of category scores)
                if score_count > 0:
                    calculated_total_score = round(total_score)
                    submission_df.at[row_index, "Total Score"] = calculated_total_score
                    logging.info(f"Calculated total score: {calculated_total_score}")
                else:
                    submission_df.at[row_index, "Total Score"] = ""
                    
                # Update the summary feedback
                submission_df.at[row_index, "Summary Feedback"] = results["summary_feedback"]
            except Exception as e:
                logging.error(f"Error processing student {student_id}: {str(e)}")
                continue
        
        # Create a backup of the original file
        original_file = CONFIG["submission_file"]
        backup_file = original_file.replace(".xlsx", f"_BACKUP_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        submission_df.to_excel(backup_file, index=False)
        logging.info(f"Backup saved to {backup_file}")
        
        # Save the updated spreadsheet
        submission_df.to_excel(original_file, index=False)
        logging.info(f"Updated spreadsheet saved to {original_file}")
        
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}", exc_info=True)

if __name__ == "__main__":
    # Run in test mode (only process first student)
    main(test_mode=True)
