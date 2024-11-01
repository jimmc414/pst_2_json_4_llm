import pypff
import argparse
import json
from datetime import datetime
import os
import sys

def parse_arguments():
    parser = argparse.ArgumentParser(description='Search PST files and extract emails based on criteria.')

    parser.add_argument('pst_path', help='Path to the PST file.')

    parser.add_argument('--date-from', type=str, help='Start date in YYYY-MM-DD format.')
    parser.add_argument('--date-to', type=str, help='End date in YYYY-MM-DD format.')
    parser.add_argument('--from', dest='from_address', type=str, help="Sender's email address.")
    parser.add_argument('--subject', type=str, help='Keyword to search in the subject.')
    parser.add_argument('--body', type=str, help='Keyword to search in the email body.')
    parser.add_argument('--output', type=str, default='output.json', help='Output JSON file path.')

    return parser.parse_args()

def open_pst(pst_path):
    if not os.path.isfile(pst_path):
        print(f"Error: The file '{pst_path}' does not exist.")
        sys.exit(1)
    pst = pypff.file()
    try:
        pst.open(pst_path)
    except Exception as e:
        print(f"Error opening PST file: {e}")
        sys.exit(1)
    return pst

def get_folder_emails(folder, criteria, emails):
    for i in range(folder.number_of_sub_folders):
        sub_folder = folder.get_sub_folder(i)
        get_folder_emails(sub_folder, criteria, emails)
    
    for i in range(folder.number_of_messages):
        message = folder.get_message(i)
        email = extract_email_fields(message)
        if email_matches_criteria(email, criteria):
            emails.append(email)

def extract_email_fields(message):
    email = {}
    email['subject'] = message.subject or ""
    email['body'] = message.plain_text_body or message.html_body or ""
    email['from'] = message.sender_name or ""
    email['to'] = message.display_to or ""
    # Attempt to parse the delivery time
    try:
        if message.delivery_time:
            email['delivery_time'] = message.delivery_time.isoformat()
        else:
            email['delivery_time'] = None
    except:
        email['delivery_time'] = None
    return email

def email_matches_criteria(email, criteria):
    # Check date range
    if criteria['date_from'] or criteria['date_to']:
        if not email['delivery_time']:
            return False
        email_date = datetime.fromisoformat(email['delivery_time'])
        if criteria['date_from'] and email_date < criteria['date_from']:
            return False
        if criteria['date_to'] and email_date > criteria['date_to']:
            return False

    # Check from address
    if criteria['from_address']:
        if criteria['from_address'].lower() not in email['from'].lower():
            return False

    # Check subject
    if criteria['subject']:
        if criteria['subject'].lower() not in email['subject'].lower():
            return False

    # Check body
    if criteria['body']:
        if criteria['body'].lower() not in email['body'].lower():
            return False

    return True

def main():
    args = parse_arguments()

    # Parse date arguments
    criteria = {}
    if args.date_from:
        try:
            criteria['date_from'] = datetime.strptime(args.date_from, '%Y-%m-%d')
        except ValueError:
            print("Error: --date-from must be in YYYY-MM-DD format.")
            sys.exit(1)
    else:
        criteria['date_from'] = None

    if args.date_to:
        try:
            criteria['date_to'] = datetime.strptime(args.date_to, '%Y-%m-%d')
        except ValueError:
            print("Error: --date-to must be in YYYY-MM-DD format.")
            sys.exit(1)
    else:
        criteria['date_to'] = None

    criteria['from_address'] = args.from_address
    criteria['subject'] = args.subject
    criteria['body'] = args.body

    # Open PST
    pst = open_pst(args.pst_path)

    # Get root folder
    root = pst.get_root_folder()

    # Collect emails
    emails = []
    get_folder_emails(root, criteria, emails)

    # Output to JSON
    try:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(emails, f, indent=4, ensure_ascii=False)
        print(f"Successfully extracted {len(emails)} emails to '{args.output}'.")
    except Exception as e:
        print(f"Error writing to JSON file: {e}")
        sys.exit(1)

    pst.close()

if __name__ == '__main__':
    main()
