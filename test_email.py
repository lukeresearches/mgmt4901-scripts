import pandas as pd
import base64
import os
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Define the Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Load feedback data
df = pd.read_csv("3A Final_Formatted_Feedback.csv")  # Replace with your actual CSV file path

# Authenticate with Gmail API using credentials.json from Google Cloud Console
def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

# Create email body from a template row
def create_email_body(row):
    return f"""Hi {row['First_Name']},

Thanks for your submission for Assignment 3A. Here’s my quick feedback for you (some of it manual, and some of it AI generated):

**Professor Feedback:**
{row['Professor Feedback']}

To support your learning, I also asked ChatGPT-4o to evaluate your work using the rubric posted on Brightspace. The feedback below was generated automatically based on your responses in specific sections of the assignment. Since this assignment was ungraded (you receive full marks for completion), I chose not to manually review each AI-generated response, instead spot checking the qualitative feedback — so please take them with a grain of salt. That said, they should still offer helpful insight and nudge your thinking forward.

**Here’s what the AI shared:**

**Capstone Execution:**
{row['Capstone Execution_feedback']}
Score (out of 20): {row['Capstone Execution_score']} 
*(Note from Luke - Don’t worry if this is low – the rubric didn’t fully apply to this assignment so it is marked rather low. You received full marks for completion.)*

**Hypothesis Testing:**
{row['Hypothesis Testing_feedback']}
Score (out of 20): {row['Hypothesis Testing_score']} 
*(Note from Luke - Don’t worry if this is low – the rubric didn’t fully apply to this assignment so it is marked rather low. You received full marks for completion.)*

**Evaluation / Decision:**
{row['Evaluation / Decision_feedback']}
Score (out of 20): {row['Evaluation / Decision_score']} 
*(Note from Luke - Don’t worry if this is low – the rubric didn’t fully apply to this assignment so it is marked rather low. You received full marks for completion.)*

Keep in mind that we didn’t cover all rubric categories (e.g., Hypothesis Development) in this assignment, so this isn’t fully representative of the next one — but you’re on the right track. Keep up the great work!

If you have any questions about the feedback or want to talk through next steps, feel free to reach out anytime.

All the best,
Luke
"""

# Create MIME message
def create_message(to, subject, body):
    message = MIMEText(body)
    message['to'] = to
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

# Send email
def send_email(service, message):
    return service.users().messages().send(userId="me", body=message).execute()

if __name__ == '__main__':
    gmail_service = authenticate_gmail()

    for i, row in df.iterrows():
        to = row['E-Mail Address']
        subject = "Your Feedback for Assignment 3A – MGMT 4901"
        body = create_email_body(row)
        msg = create_message(to, subject, body)
        print(f"Sending to {to}...")
        send_email(gmail_service, msg)
        print("Sent!")
