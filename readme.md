## Overview

1. **Dependencies**:
    - **Python 3.6+**
    - **`pypff` Library**: To read PST files.
    - **`argparse`**: To parse command-line arguments.
    - **`json`**: To output the results in JSON format.
    - **`datetime`**: To handle date-related operations.

2. **Features**:
    - Accepts the path to a PST file.
    - Allows filtering based on:
        - Date range (`--date-from` and `--date-to`).
        - Sender's email address (`--from`).
        - Subject content (`--subject`).
        - Body content (`--body`).
    - Outputs the matching emails to a JSON file.

## Prerequisites

### 1. Install Python

Ensure you have Python 3.6 or later installed. You can download it from the [official website](https://www.python.org/downloads/).

### 2. Install `pypff`

The `pypff` library is a Python binding for the [libpff](https://github.com/libyal/libpff) library, which allows you to read PST files.

#### Installation Steps:

**For Windows:**

1. **Download Precompiled Binaries**:

   - Visit the [libpff releases page](https://github.com/libyal/libpff/releases).
   - Download the latest Windows binaries.
   - Extract the files to a directory, e.g., `C:\libpff`.

2. **Install `pypff`**:

   - Open Command Prompt.
   - Navigate to the `pypff` directory within the extracted files.
   - Run the following command:
     ```bash
     python setup.py install
     ```

**For macOS/Linux:**

1. **Install Dependencies**:

   - Ensure you have `libpff` installed. You can install it using `brew` (for macOS) or `apt`/`yum` (for Linux).

   ```bash
   # macOS
   brew install libpff

   # Ubuntu/Debian
   sudo apt-get update
   sudo apt-get install libpff-dev python3-pip

   # Fedora
   sudo dnf install libpff-devel python3-pip
   ```

2. **Install `pypff`**:

   ```bash
   pip3 install pypff
   ```

*Note*: Installing `pypff` can be challenging due to its dependencies. If you encounter issues, refer to the [libpff documentation](https://github.com/libyal/libpff) for detailed installation instructions.

### 3. Install Other Python Dependencies

Open your terminal or command prompt and install the required Python libraries:

```bash
pip install argparse
```

*Note*: The `argparse`, `json`, and `datetime` modules are part of Python's standard library and do not require separate installation.

## The Python Program

```python
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
```

## How the Program Works

1. **Argument Parsing**: The program uses `argparse` to handle command-line arguments, allowing users to specify the PST file path and various search criteria.

2. **Opening the PST File**: Utilizes `pypff` to open and read the PST file.

3. **Traversing Folders**: Recursively navigates through all folders and subfolders in the PST to access all emails.

4. **Extracting Email Fields**: For each email, it extracts relevant fields such as subject, body, sender, recipients, and delivery time.

5. **Applying Search Criteria**: Filters emails based on the provided criteria:
    - **Date Range**: Filters emails within the specified start and end dates.
    - **Sender**: Filters emails from a specific sender.
    - **Subject**: Searches for keywords in the subject.
    - **Body**: Searches for keywords in the email body.

6. **Outputting to JSON**: Writes the filtered emails to a JSON file in a structured format.

## Running the Program

1. **Save the Script**:

   Save the provided Python code to a file, e.g., `pst_search.py`.

2. **Execute the Script**:

   Open your terminal or command prompt and navigate to the directory containing `pst_search.py`.

3. **Basic Usage**:

   ```bash
   python pst_search.py path_to_pst_file.pst --output result.json
   ```

4. **With Search Criteria**:

   - **Filter by Date Range**:

     ```bash
     python pst_search.py path_to_pst_file.pst --date-from 2023-01-01 --date-to 2023-12-31 --output result.json
     ```

   - **Filter by Sender**:

     ```bash
     python pst_search.py path_to_pst_file.pst --from sender@example.com --output result.json
     ```

   - **Filter by Subject and Body**:

     ```bash
     python pst_search.py path_to_pst_file.pst --subject "Meeting" --body "Project Update" --output result.json
     ```

   - **Combine Multiple Criteria**:

     ```bash
     python pst_search.py path_to_pst_file.pst --date-from 2023-01-01 --date-to 2023-12-31 --from sender@example.com --subject "Meeting" --body "Project Update" --output result.json
     ```

5. **Example**:

   ```bash
   python pst_search.py "C:\Users\YourName\Documents\emails.pst" --date-from 2023-01-01 --date-to 2023-06-30 --from "john.doe@example.com" --subject "Invoice" --output filtered_emails.json
   ```

   This command searches for emails in `emails.pst` sent by `john.doe@example.com` with "Invoice" in the subject, received between January 1, 2023, and June 30, 2023, and outputs the results to `filtered_emails.json`.

## JSON Output Structure

The output JSON file will be an array of email objects with the following structure:

```json
[
    {
        "subject": "Meeting Agenda",
        "body": "Dear team,\nPlease find the agenda for tomorrow's meeting...",
        "from": "john.doe@example.com",
        "to": "jane.smith@example.com; mark.brown@example.com",
        "delivery_time": "2023-03-15T10:30:00"
    },
    {
        "subject": "Project Update",
        "body": "The project is on track for the Q2 release...",
        "from": "jane.smith@example.com",
        "to": "john.doe@example.com",
        "delivery_time": "2023-04-20T14:45:00"
    }
    // More emails...
]
```

- **subject**: The subject line of the email.
- **body**: The plain text or HTML body of the email.
- **from**: The sender's name or email address.
- **to**: The recipients' names or email addresses.
- **delivery_time**: The date and time the email was delivered in ISO 8601 format.

## Notes and Tips

- **Performance**: Parsing large PST files can be time-consuming. Ensure you have sufficient system resources and consider running the script on a machine with adequate performance.

- **Encoding Issues**: Some emails might contain characters that need proper encoding. The script uses `utf-8` encoding to handle most cases.

- **Error Handling**: The script includes basic error handling for file operations and PST parsing. For more robust applications, consider enhancing error handling mechanisms.

- **Extensibility**: You can extend the script to include additional search criteria or output formats as needed.

## Troubleshooting

1. **`pypff` Installation Issues**:

   - Ensure all dependencies are installed.
   - Refer to the [libpff GitHub repository](https://github.com/libyal/libpff) for detailed installation guides and troubleshooting tips.
   - Consider using Docker or virtual environments to manage dependencies.

2. **PST File Access Issues**:

   - Ensure the PST file is not corrupted.
   - Verify that you have the necessary permissions to read the PST file.

3. **Script Errors**:

   - Ensure all command-line arguments are correctly specified.
   - Check for typos in email addresses or search keywords.
   - Validate date formats (`YYYY-MM-DD`).

## Conclusion

This Python script provides a flexible way to search and extract emails from PST files based on various criteria, outputting the results in a structured JSON format. With proper setup and dependencies, it can be a powerful tool for managing and analyzing your email data.