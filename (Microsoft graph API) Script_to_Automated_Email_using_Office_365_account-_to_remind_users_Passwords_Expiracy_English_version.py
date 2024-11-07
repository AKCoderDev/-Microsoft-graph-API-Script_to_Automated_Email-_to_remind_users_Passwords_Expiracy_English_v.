import msal
import requests
from datetime import datetime, timedelta, timezone
from ldap3 import Server, Connection, ALL
import os
import sys

# Microsoft Graph API configuration
client_id = 'ID'            # Application (client) ID
tenant_id = 'ID'            # Directory (tenant) ID
client_secret = 'key'       # Client secret
user_email = 'it_team@example.com'   # Email of the user sending the notification

# MSAL settings for obtaining an access token
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

# Notification settings
expire_in_days = 5  # Number of days before password expiration to notify the user
logging_enabled = True  # Enable logging
log_file = os.path.join(os.getcwd(), f"PasswordChangeNotification_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.csv")  # Log file
testing_enabled = True  # Enable/disable test mode (True/False)
test_recipient = "recipient@example.com"  # Test recipient email address

# Class to capture print() output for logging
class TeeOutput:
    def __init__(self, file_path):
        self.file = open(file_path, 'a', encoding='utf-8', buffering=1)
        self.stdout = sys.stdout

    def write(self, message):
        self.file.write(message)
        self.stdout.write(message)

    def flush(self):
        self.file.flush()
        self.stdout.flush()

sys.stdout = TeeOutput(log_file)

# Function to obtain an access token via MSAL
def get_access_token():
    app = msal.ConfidentialClientApplication(
        client_id,
        client_credential=client_secret,
        authority=authority
    )
    result = app.acquire_token_for_client(scopes=scope)
    if "access_token" in result:
        return result["access_token"]
    else:
        print("Failed to obtain access token:", result.get("error"), result.get("error_description"))
        return None

# Function to send email via Microsoft Graph API
def send_email_via_graph_api(to_email, subject, body, access_token):
    email_msg = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email
                    }
                }
            ]
        }
    }
    graph_api_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(graph_api_url, headers=headers, json=email_msg)

    # Log the API response
    print("API Response:", response.status_code, response.text)  # Logging API response

    if response.status_code == 202:
        print(f"Email successfully sent to {to_email}")
    else:
        print("Error sending email:", response.status_code, response.text)

# Function for logging events to a CSV file
def log_event(log_message):
    if logging_enabled:
        try:
            with open(log_file, 'a') as f:
                f.write(log_message + "\n")
            print("Saved to file:", log_message)  # Console logging
        except Exception as e:
            print(f"Error logging to file: {e}")  # Error handling

# Function to retrieve users from Active Directory
def get_ad_users():
    server = Server('ldap://domain.local', get_info=ALL)
    conn = Connection(server, 'account@domain.local', 'password', auto_bind=True)
    conn.search(
        'dc=domain,dc=local',
        '(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(mail=*)(pwdLastSet>=0))',
        attributes=['cn', 'mail', 'pwdLastSet'])
    users = conn.entries
    return users

# Main function to process users and send password expiration notifications
def process_users():
    access_token = get_access_token()
    if not access_token:
        print("No access token, ending process.")
        return

    users = get_ad_users()
    print(f"Found {len(users)} users.")  # Added logging for the number of users
    for user in users:
        name = str(user.cn)
        email = str(user.mail)
        pwd_last_set = user.pwdLastSet.value  # Using 'int' for date as an integer
        max_password_age = timedelta(days=90)
        expires_on = pwd_last_set + max_password_age
        days_to_expire = (expires_on - datetime.now(timezone.utc)).days
        print(f"User: {name}, email: {email}, days_to_expire: {days_to_expire}")
        print(f"pwd_last_set for {name}: {pwd_last_set}")
        print(f"expires_on for {name}: {expires_on}")

        if days_to_expire <= expire_in_days and days_to_expire >= 0:
            message_days = f"in {days_to_expire} days" if days_to_expire > 0 else "today"
            subject = f"!TEST! Your password will expire {message_days} !TEST!"
            body = f"""
            <p>Dear {name},<br></p>
            <p> !TEST! Your password will expire {message_days}. Please remember to change it.<br>
            For help, read this article: <a href='https://link/to/article/'>How to change your password?</a><br>
            Or contact Helpdesk: <a href='mailto:helpdesk@example.com'>helpdesk@example.com  !TEST! </a><br></p>
            <p>Best regards,<br>IT Helpdesk</p>
            """
            print("Checking days to expire")
            recipient_email = test_recipient if testing_enabled else email  # Configuring test recipient
            
            if not email:
                email = "it_team@example.com"

            log_event(
                f"Date,Name,Email,DaysToExpire,ExpiresOn\n"
                f"{datetime.now().strftime('%Y-%m-%d')},{name},{email},{days_to_expire},{expires_on}")

            print(f"Attempting to send email to: {email}, Subject: {subject}")  # Logging before sending
            send_email_via_graph_api(recipient_email, subject, body, access_token)

# Run the main function
if __name__ == "__main__":
    process_users()
input("The end")