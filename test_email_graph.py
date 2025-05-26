import requests
import json
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Microsoft Graph API credentials are no longer used
email_address = "lk701947@dal.ca"

def get_access_token():
    """Get access token from Microsoft Graph API."""
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    
    payload = {
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials'
    }
    
    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        return response.json()['access_token']
    except Exception as e:
        logging.error(f"Error getting access token: {str(e)}")
        raise

def send_test_email():
    try:
        # Get access token
        access_token = get_access_token()
        
        # Email content
        subject = "✅ Test Email from Python (MGMT 4901)"
        body = """Hi there,

This is a test to confirm that Python can send emails from my Microsoft university email account.

If you're seeing this — it worked!

Best,
Luke (via Python)"""
        
        # Recipient and sender
        sender = os.getenv('EMAIL_ADDRESS')
        recipient = "luke.decoste@cgu.edu"
        
        # Create email message
        message = {
            'message': {
                'subject': subject,
                'body': {
                    'contentType': 'text',
                    'content': body
                },
                'toRecipients': [
                    {
                        'emailAddress': {
                            'address': recipient
                        }
                    }
                ]
            },
            'saveToSentItems': 'true'
        }
        
        # Send email using Graph API
        graph_url = 'https://graph.microsoft.com/v1.0/users/me/sendMail'
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            response = requests.post(graph_url, headers=headers, json=message)
            response.raise_for_status()
            logging.info("✅ Email sent successfully!")
        except Exception as e:
            logging.error(f"Error sending email: {str(e)}")
            raise
            
    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")
        raise

if __name__ == "__main__":
    send_test_email()
