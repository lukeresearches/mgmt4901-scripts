import pandas as pd
import base64
import os
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import chardet # Added for encoding detection
import os # Ensure os is imported for path operations

# Define the Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Load feedback data
file_path = "/Users/decosteluke/Dropbox/ACademic  Teaching - Dalhousie/2025-05 - MGMT 4901 Async/Assignments for Mailing/MGMT4901_3B_Evaluation_OutputR1.csv"

# Try to detect encoding first
try:
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    detected_encoding = result['encoding']
    if detected_encoding:
        print(f"Detected encoding: {detected_encoding}")
        df = pd.read_csv(file_path, encoding=detected_encoding)
    else:
        # If chardet fails to detect, try fallback
        raise ValueError("Chardet could not detect encoding.")
except Exception as e:
    print(f"Encoding detection/read failed: {e}. Trying fallback encodings...")
    fallback_encodings = ['latin1', 'ISO-8859-1', 'cp1252']
    df = None
    for enc in fallback_encodings:
        try:
            df = pd.read_csv(file_path, encoding=enc)
            print(f"Successfully read file with fallback encoding: {enc}")
            break
        except UnicodeDecodeError:
            print(f"Failed to read with encoding: {enc}")
        except Exception as ex:
            print(f"An unexpected error occurred with encoding {enc}: {ex}")
            break 
    if df is None:
        print("Failed to read the CSV file with all attempted encodings. Please check the file.")
        # Handle the error as appropriate for your script, e.g., exit or raise
        exit()


# Authenticate with Gmail API using credentials.json from Google Cloud Console
def authenticate_gmail():
    # Define the directory where credentials.json and token.json are stored
    credentials_dir = "/Users/decosteluke/Dropbox/ACademic  Teaching - Dalhousie/2025-05 - MGMT 4901 Async/"
    credentials_path = os.path.join(credentials_dir, 'credentials.json')
    token_path = os.path.join(credentials_dir, 'token.json')

    creds = None
    if os.path.exists(token_path):
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                print(f"Error: credentials.json not found at {credentials_path}")
                print(f"Please ensure 'credentials.json' is in the directory: {credentials_dir}")
                exit()
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the token to the same directory as credentials.json
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

# Create email body from a template row
def create_email_body(row):
    return f"""Hi {row['First_Name']},

Thanks for your submission for Assignment 3B. Here’s my quick feedback for you. Please review it with the feedback of your peers to align on the go-forward plan.

If you want to see your submission, it's available here: https://dalu.sharepoint.com/:x:/t/MGMTComm4901-Summer2025/EUd5L8grklZArRQjyvitGpgBzz2nqODxYDWtbsgH3BEUMQ?e=yteQFL


**Professor Feedback:**
{row['Professor Feedback']}

Scores in individual sections are below: 
**Capstone Execution:**
Score (out of 20): {row['Capstone Execution_score']} 

**Hypothesis Development:**
Score (out of 20): {row['Hypothesis Development_score']}

**Hypothesis Testing:**
Score (out of 20): {row['Hypothesis Testing_score']} 

**Evaluation / Decision:**
Score (out of 20): {row['Evaluation / Decision_score']} 


If you have any questions about the feedback or want to talk through next steps, feel free to reach out anytime.

All the best,
Luke
"""

# Create MIME message
def create_message(to, subject, body):
    message = MIMEText(body, 'html') # Ensure body is treated as HTML
    message['to'] = to
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

# Send email
def send_email(service, message):
    return service.users().messages().send(userId="me", body=message).execute()

if __name__ == '__main__':
    gmail_service = authenticate_gmail()

    # Just run a single test row for now
    row = df.iloc[0]
    to = row['E-Mail Address']
    subject = "Your Feedback for Assignment 3B – MGMT 4901"
    body = create_email_body(row)
    msg = create_message(to, subject, body)
    print(f"Sending to {to}...")
    send_email(gmail_service, msg)
    print("Sent!")
