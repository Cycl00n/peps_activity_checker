PEPS Activity Checker is a desktop application designed to help educators and administrators verify activity encodings, ensure correct participation notes, detect missing information, and quickly send reminder emails to the appropriate staff.
The application processes a PEPS Excel file, analyzes each activity, detects errors or missing data, and provides tools to automatically generate reminder emails.
The tool is optimized for use in social/educational institutions using PEPS-like activity logs.

Features:

- Reads PEPS activity files (.xlsx)

- Detects:

  - Missing general descriptions

  - Residents without notes

  - Activities without participation

  - Cancelled activities

  - Activities with multiple educators

- Automatically stops scanning future-dated entries

- Multi-Educator Support:

  - Extracts all educators involved in an activity

  - Cleans educator names from the activity title

  - Lets you choose one educator from a dropdown

  - Sends one email per educator

- Email Generation

  - Auto-fills:

    - Recipient address (based on employees.json)

    - CC field

    - Subject

    - Body (formatted nicely)

  - A green checkmark appears after sending

  - Smart Name Removal

- JSON Editing

  - Quick access to employees.json and residents.json (removed for git for privacy reasons) 
