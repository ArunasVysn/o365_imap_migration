import threading
import json
import time
import os
from O365 import Account, FileSystemTokenBackend
import imaplib
import requests
import shutil
import html
import re
import csv
import base64

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# Load OAuth2 constants from config.json
def load_config(config_file='config.json'):
    """Load configuration values from a JSON file."""
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"Configuration file {config_file} not found.")

    with open(config_file, 'r') as f:
        config = json.load(f)

    return config


# Load the configuration at the start of the program
config = load_config()

# Use the loaded configuration values
SOURCE_CLIENT_ID = config.get('SOURCE_CLIENT_ID')
SOURCE_CLIENT_SECRET = config.get('SOURCE_CLIENT_SECRET')
SOURCE_TENANT_ID = config.get('SOURCE_TENANT_ID')
MIGRATE_ATTACHMENTS = config.get('MIGRATE_ATTACHMENTS', False)

# Global variable to store the authorization code
auth_code = None

# Create a lock for thread-safe printing
print_lock = threading.Lock()

def safe_print(message, end="\n"):
    """Thread-safe print function that flushes the output immediately."""
    with print_lock:
        print(message, end=end, flush=True)  # Allow control over end character

def update_progress(message, final=False):
    """Updates progress by safely printing a message."""
    if final:
        safe_print(message)  # Final message prints with newline
    else:
        # Print with carriage return to overwrite the line
        safe_print(message, end="\r")

def authenticate_account(target_email):
    tokens_dir = 'tokens'
    account_tokens_dir = os.path.join(tokens_dir, f'token_{target_email}')
    token_path = os.path.join(account_tokens_dir, 'o365_token.txt')

    # Ensure the directory exists
    if os.path.exists(token_path):
        shutil.rmtree(account_tokens_dir)

    if not os.path.exists(account_tokens_dir):
        os.makedirs(account_tokens_dir)

    # Re-authenticate if the token doesn't exist or is invalid
    account = Account(
        (SOURCE_CLIENT_ID, SOURCE_CLIENT_SECRET),
        tenant_id=SOURCE_TENANT_ID,
        token_backend=FileSystemTokenBackend(token_path=account_tokens_dir),
        auth_flow_type='credentials'
    )

    if not account.is_authenticated:
        print(f"Starting authentication for {target_email}.")
        try:
            account.authenticate(scopes=['https://graph.microsoft.com/.default'])
            update_progress(f"Successfully authenticated for {target_email}.")

            # Now read the access token directly from the token file after authentication
            with open(token_path, 'r') as token_file:
                token_data = json.load(token_file)
                access_token = token_data.get('access_token')

            if access_token:
                return access_token, account  # Return both the access token and account object
            else:
                print("Failed to retrieve access token after authentication.")
                return None, None

        except Exception as e:
            print(f"An error occurred during authentication for {target_email}: {e}")
            return None, None

    return None, account  # If already authenticated, return the account object without the access token


def fetch_attachments(email_address, message_id, access_token):
    """Fetch attachments for a given message."""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages/{message_id}/attachments"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            attachments = response.json().get('value', [])
            return attachments
        else:
            print(f"Failed to fetch attachments for message {message_id}: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        print(f"Error fetching attachments for message {message_id}: {e}")
        return []


def fetch_all_emails(account, source_email, access_token):
    """Fetch all emails from the source email account."""
    if not account.is_authenticated:
        print("Account is not authenticated. Cannot fetch emails.")
        return []

    mailbox = account.mailbox()
    folders = get_mail_folders(access_token, source_email)
    if folders is None:
        print("Failed to retrieve folders.")
        return []

    all_emails = []

    for folder in folders.get('value', []):
        try:
            folder_id = folder['id']
            url = f"https://graph.microsoft.com/v1.0/users/{source_email}/mailFolders/{folder_id}/messages"
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }

            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                messages = response.json().get('value', [])
                for message in messages:
                    message['folderName'] = folder['displayName']  # Store the folder name with the message

                    # Fetch attachments if enabled in config
                    if MIGRATE_ATTACHMENTS:
                        message['attachments'] = fetch_attachments(source_email, message['id'], access_token)

                all_emails.extend(messages)
                update_progress(f"Fetched {len(messages)} emails from {folder['displayName']}.")
            else:
                print(f"Failed to fetch emails from folder {folder['displayName']}: {response.status_code} - {response.text}")

        except Exception as e:
            print(f"Error fetching emails from folder {folder['displayName']}: {e}")

    return all_emails


def get_mail_folders(access_token, user_email):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        folders = response.json()
        return folders  # Successfully retrieved folders
    else:
        print(f"Error fetching mail folders: {response.status_code} - {response.text}")
        return None


def extract_email_body(msg):
    """Extracts and cleans the email body from the message."""
    body = ""

    if 'body' in msg and 'content' in msg['body']:
        if msg['body']['contentType'] == 'html':
            html_content = msg['body']['content']
            body = re.sub(r'<[^>]+>', '', html_content)  # Remove tags
            body = re.sub(r'\s+', ' ', body)  # Replace multiple spaces/newlines with a single space
        elif msg['body']['contentType'] == 'text':
            body = msg['body']['content']

    body = html.unescape(body.strip())

    return body if body else "(No Body)"  # Default if no body found


def convert_to_rfc822(msg):
    """Converts a message dictionary to RFC822 format with attachments, only if necessary."""
    try:
        subject = msg.get('subject', '(No Subject)')
        sender = msg.get('from', {}).get('emailAddress', {}).get('address', 'unknown@example.com')

        recipients = msg.get('toRecipients', [])
        recipient_list = ', '.join(
            [rec.get('emailAddress', {}).get('address', 'unknown@example.com') for rec in recipients])

        cc_recipients = msg.get('ccRecipients', [])
        cc_list = ', '.join(
            [rec.get('emailAddress', {}).get('address', 'unknown@example.com') for rec in cc_recipients])

        bcc_recipients = msg.get('bccRecipients', [])
        bcc_list = ', '.join(
            [rec.get('emailAddress', {}).get('address', 'unknown@example.com') for rec in bcc_recipients])

        all_recipients = recipient_list
        if cc_list:
            all_recipients += f", {cc_list}"
        if bcc_list:
            all_recipients += f", {bcc_list}"

        date = msg.get('receivedDateTime', 'Tue, 01 Jan 2000 00:00:00 +0000')
        body = extract_email_body(msg)

        if not isinstance(body, str):
            print(f"Body is not a string for message: {msg}")  # Log the problematic message
            body = "(Invalid Body)"  # Default if body is not a string

        # Determine if the message has attachments
        has_attachments = MIGRATE_ATTACHMENTS and 'attachments' in msg and msg['attachments']

        if has_attachments:
            # Create a multipart email message when there are attachments
            email_msg = MIMEMultipart()
            email_msg.attach(MIMEText(body, 'plain'))  # Attach the body as plain text

            # Process and attach each attachment
            for attachment in msg['attachments']:
                try:
                    mime_attachment = MIMEBase('application', 'octet-stream')
                    attachment_content = base64.b64decode(attachment.get('contentBytes', ''))

                    mime_attachment.set_payload(attachment_content)
                    encoders.encode_base64(mime_attachment)

                    mime_attachment.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{attachment.get("name", "unknown")}"'
                    )

                    email_msg.attach(mime_attachment)
                except Exception as e:
                    print(f"Error processing attachment: {e}")
        else:
            # Create a plain text message if no attachments
            email_msg = MIMEText(body, 'plain')

        # Set email headers (From, To, Subject, Date)
        email_msg['From'] = sender
        email_msg['To'] = all_recipients
        email_msg['Subject'] = subject
        email_msg['Date'] = date

        # Return the email message in RFC822 format
        return email_msg.as_string()

    except Exception as e:
        print(f"Conversion error: {e}, message: {msg}")  # Log the error and message
        return None  # Return None if conversion fails



def connect_to_target_imap(server, email_user, password=None):
    try:
        mail = imaplib.IMAP4_SSL(server, timeout=10)
        mail.login(email_user, password)
        print(f"Successfully connected to {server} as {email_user}.")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Failed to connect to {server}: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

def select_target_folder(target_mail, source_folder_name):
    """Selects a target folder based on the source folder name."""
    try:
        # Get a list of available folders in the target mailbox
        status, folders = target_mail.list()
        available_folders = [folder.decode().split(' "/" ')[-1] for folder in folders]

        # Log available folders
        print(f"Available target folders: {available_folders}")

        # Clean the folder name for matching
        target_folder_name = clean_folder_name(source_folder_name)

        # Check if the target folder exists, otherwise return 'INBOX'
        if target_folder_name in available_folders:
            return target_folder_name
        else:
            print(f"Target folder '{target_folder_name}' not found. Using 'INBOX'.")
            return 'INBOX'
    except Exception as e:
        print(f"Error fetching folders: {e}")
        return 'INBOX'  # Default to INBOX on error


def get_target_folders(target_mail):
    """Fetches the available folders in the target mail account and returns a mapping."""
    folder_mapping = {
        "Inbox": "INBOX",
        "Drafts": "Drafts",
        "Sent Items": "Sent",
        "Deleted Items": "Deleted Items",  # Update if your target IMAP server uses a different name
        "Junk Email": "Spam"  # Update if your target IMAP server uses a different name
    }
    return folder_mapping


def migrate_emails(target_mail, messages):
    """Migrates emails from a list of message dictionaries to a target email account."""
    target_folders = get_target_folders(target_mail)  # Get the mapping of target folders

    total_messages = len(messages)  # Total number of messages
    for idx, msg in enumerate(messages):  # Iterate with index for progress tracking
        try:
            if isinstance(msg, dict):
                subject = msg.get('subject', '(No Subject)')
                source_folder = msg.get('folderName', 'Inbox')

                # Determine target folder name based on the source folder
                target_folder_name = target_folders.get(source_folder, "INBOX")

                # Convert email message to RFC822 format
                email_message = convert_to_rfc822(msg)

                if isinstance(email_message, str):
                    email_message_bytes = email_message.encode('utf-8')
                    retry_count = 3
                    for attempt in range(retry_count):
                        try:
                            # Append email message to the target folder
                            status, response = target_mail.append(
                                target_folder_name, None,
                                imaplib.Time2Internaldate(time.time()), email_message_bytes
                            )
                            if status != 'OK':
                                update_progress(f"Failed to append message to {target_folder_name}: {response}")
                            break  # Exit retry loop on success
                        except (OSError, imaplib.IMAP4.abort) as e:
                            time.sleep(2)  # Wait before retrying
                        except Exception as e:
                            update_progress(f"Error appending message: {e}")  # Log the error
                            break
                else:
                    update_progress(f"Email conversion failed for message: {msg}")
            else:
                update_progress(f"Unexpected message format: {msg}")

            # Custom progress update
            progress_percentage = (idx + 1) / total_messages * 100
            update_progress(f"Migrating Emails: {progress_percentage:.2f}% ({idx + 1}/{total_messages})")

        except Exception as e:
            update_progress(f"Error migrating email: {e}")


def clean_folder_name(folder_name):
    """Cleans the folder name for compatibility with the target IMAP server."""
    folder_name = re.sub(r'[<>:"/\\|?*]', '', folder_name)  # Remove illegal characters
    return folder_name.strip()[:50]  # Trim to 50 characters if needed

def read_mailboxes_from_csv(filename):
    """Reads mailboxes from a CSV file."""
    mailboxes = []
    try:
        with open(filename, 'r') as csvfile:
            csvreader = csv.reader(csvfile)
            next(csvreader)  # Skip the header
            for row in csvreader:
                if len(row) == 4:  # Ensure the CSV has four columns
                    source_email = row[0]
                    target_server = row[1]
                    target_email = row[2]
                    target_password = row[3]
                    mailboxes.append((source_email, target_server, target_email, target_password))
    except Exception as e:
        print(f"Error reading CSV file: {e}")
    return mailboxes


def migrate_mailbox(mailbox):
    """Migrates a single mailbox."""
    source_email, target_server, target_email, target_password = mailbox  # Unpack all four columns

    update_progress(f"Starting migration from {source_email} to {target_email}.")

    access_token, account = authenticate_account(source_email)
    if access_token is None:
        update_progress(f"Authentication failed for {source_email}. Skipping this mailbox.")
        return

    messages = fetch_all_emails(account, source_email, access_token)

    if messages:
        # Connect to target IMAP server with the specified server and password
        target_mail = connect_to_target_imap(target_server, target_email, target_password)  # Use the target server and password
        if target_mail is None:
            update_progress(f"Failed to connect to target IMAP server for {target_email}.")
            return

        migrate_emails(target_mail, messages)
        target_mail.logout()  # Log out from the target IMAP server
        update_progress(f"Migration completed for {source_email} to {target_email}.")
        print()  # Move to next line after migration completion
    else:
        update_progress(f"No messages to migrate from {source_email}.")



def main():
    csv_filename = 'details.csv'  # Replace with your CSV filename
    mailboxes = read_mailboxes_from_csv(csv_filename)
    if not mailboxes:
        print("No mailboxes found to migrate.")
        return

    threads = []
    for mailbox in mailboxes:
        thread = threading.Thread(target=migrate_mailbox, args=(mailbox,))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    print("All mailboxes have been migrated.")

if __name__ == "__main__":
    main()
