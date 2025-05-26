import pandas as pd
import openai
from dotenv import load_dotenv
import os
import json
import logging
from typing import Dict, List, Tuple

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuration header
CONFIG = {
    "rubric_file": "2025_05_Rubric_Table.xlsx",
    "rubric_sheet": "Rubric",
    "submission_file": "4) Initial Belief Formation  (Describe 1) Initial Theory of Value and 2) supporting hypotheses.  - Attempt Details_CLEANED.xlsx",
    "id_column": "username",
    "output_csv": "MGMT4901_Evaluation_Output.csv"
}

def load_rubric() -> Dict[str, List[Dict]]:
    """Load rubric from Excel file and organize by category."""
    try:
        rubric_df = pd.read_excel(CONFIG["rubric_file"], sheet_name=CONFIG["rubric_sheet"])
        rubric_by_category = {}
        
        for _, row in rubric_df.iterrows():
            category = row['Rubric Category']
            if category not in rubric_by_category:
                rubric_by_category[category] = []
            rubric_by_category[category].append({
                'id': row['ID'],
                'item': row['Rubric Item'],
                'weighting': row['Weighting']
            })
        
        return rubric_by_category
    except Exception as e:
        logging.error(f"Error loading rubric: {str(e)}")
        raise

def collect_category_responses(student_responses: Dict, category: str) -> str:
    """Collect all responses for a given rubric category."""
    responses = []
    for col, response in student_responses.items():
        if isinstance(response, str) and category in col:
            responses.append(f"{col}: {response}")
    return "\n".join(responses)

def create_prompt(category: str, rubric_items: List[Dict], student_responses: str) -> str:
    """Create the evaluation prompt for OpenAI API."""
    rubric_items_text = "\n".join([f"- {item['item']}" for item in rubric_items])
    
    return f"""You are acting as the instructor for MGMT 4901: Applied Entrepreneurship & Innovation. You are evaluating one team's submission using a structured rubric.

You will review the student's written responses (provided below) and evaluate them using a specific rubric category that contains four rubric items. Each item is worth 5%, for a total of 20% for the category.

Your task is to:
1. Write a 1–2 sentence comment on the student's performance in this category
2. Provide a numeric score out of 20 (or leave blank if insufficient data)

Use the following voice and tone guidelines:

---
**Luke DeCoste Feedback Style Guide**  
- Warm and affirming, but never indulgent. Be grounded and respectful.  
- Encourage student agency and reflection — don't fix their work for them.  
- Start with what's working. Then ask a thoughtful question or suggest a next step.  
- Avoid overenthusiastic or canned praise. Be specific and plainspoken.  
- Use active voice and curious phrasing, e.g.:
  - "What's emerging here is promising..."
  - "You might consider..."
  - "One thing I'm wondering is..."

---

### Rubric Category: {category}

Rubric Items:
{rubric_items_text}

Student Responses:
{student_responses}

Respond in JSON with the following format:

```json
{{
  "rubric_category": "{category}",
  "feedback": "Your feedback here",
  "score": score_value
}}
```
If no relevant answers were provided, still write a thoughtful reflection prompt but leave "score": "".
"""

def evaluate_category(student_id: str, category: str, rubric_items: List[Dict], student_responses: Dict) -> Dict:
    """Evaluate a single rubric category for a student."""
    responses_text = collect_category_responses(student_responses, category)
    
    if not responses_text.strip():
        return {
            "rubric_category": category,
            "feedback": f"No responses were provided for the {category} category. Consider reflecting on how this aspect could be strengthened in your submission.",
            "score": ""
        }
    
    prompt = create_prompt(category, rubric_items, responses_text)
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful teaching assistant. Please respond ONLY with a valid JSON object in this exact format: {\"rubric_category\":\"...\", \"feedback\":\"...\", \"score\":...}. Do not include any other text or explanations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        
        # Extract JSON from response
        try:
            # Clean up the response content
            content = response.choices[0].message.content.strip()
            
            # Try to extract JSON if it's in a code block
            if content.startswith("```json") and content.endswith("```"):
                content = content[7:-3].strip()  # Remove code block markers
            
            # If we have a string, try to add quotes around it
            if content and not content.startswith("{"):
                content = f'{{"rubric_category": "{category}", "feedback": "{content}", "score": ""}}'
            
            # Try to parse the JSON
            evaluation = json.loads(content)
            
            # Validate the structure
            if not isinstance(evaluation, dict) or \
               "rubric_category" not in evaluation or \
               "feedback" not in evaluation or \
               "score" not in evaluation:
                raise ValueError("Invalid JSON structure")
                
            return evaluation
            
        except (json.JSONDecodeError, ValueError) as e:
            logging.error(f"Invalid JSON response for student {student_id}, category {category}: {str(e)}")
            # Try to create a fallback evaluation
            try:
                # Try to extract just the feedback text
                feedback = content.strip()
                if not feedback:
                    feedback = "Error processing evaluation. Please try again."
                
                return {
                    "rubric_category": category,
                    "feedback": feedback,
                    "score": ""
                }
            except:
                return {
                    "rubric_category": category,
                    "feedback": "Error processing evaluation. Please try again.",
                    "score": ""
                }
    except Exception as e:
        logging.error(f"API error for student {student_id}, category {category}: {str(e)}")
        return {
            "rubric_category": category,
            "feedback": "Error processing evaluation. Please try again.",
            "score": ""
        }

def evaluate_all_categories(student_id: str, rubric_by_category: Dict, submission_df: pd.DataFrame) -> List[Dict]:
    """Evaluate all rubric categories for a single student."""
    student_responses = submission_df[submission_df[CONFIG["id_column"]] == student_id].iloc[0].to_dict()
    evaluations = []
    
    for category, rubric_items in rubric_by_category.items():
        evaluation = evaluate_category(student_id, category, rubric_items, student_responses)
        evaluations.append({
            "username": student_id,
            "rubric_category": evaluation["rubric_category"],
            "feedback": evaluation["feedback"],
            "score": evaluation["score"]
        })
    
    return evaluations

def main():
    """Main function to process all submissions."""
    # Load OpenAI API key
    load_dotenv()
    openai.api_key = os.getenv('OPENAI_API_KEY')
    
    try:
        # Load rubric
        rubric_by_category = load_rubric()
        
        # Load submission file
        submission_df = pd.read_excel(CONFIG["submission_file"])
        
        # Get unique usernames
        usernames = submission_df[CONFIG["id_column"]].unique()
        
        # Process each student
        all_evaluations = []
        for username in usernames:
            logging.info(f"Evaluating submissions for user: {username}")
            evaluations = evaluate_all_categories(username, rubric_by_category, submission_df)
            all_evaluations.extend(evaluations)
        
        # Create output DataFrame
        output_df = pd.DataFrame(all_evaluations)
        
        # Save to CSV
        output_df.to_csv(CONFIG["output_csv"], index=False)
        logging.info(f"Evaluation complete. Results saved to {CONFIG["output_csv"]}")
        
    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")
        raise

if __name__ == "__main__":
    main()
