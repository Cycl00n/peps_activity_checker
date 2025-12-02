import platform
import sys


def send_email_outlook(to, cc, subject, body):
    """
    Send email through Outlook on Windows

    Args:
        to (str): Recipient email address
        cc (str): CC email addresses (comma-separated)
        subject (str): Email subject
        body (str): Email body text

    Returns:
        tuple: (success: bool, message: str)
    """

    if platform.system() != "Windows":
        return False, "Outlook integration only works on Windows"

    try:
        import win32com.client
    except ImportError:
        return (
            False,
            "pywin32 not installed.\n\nRun: pip install pywin32\n\nThen restart the program.",
        )

    try:
        # Create Outlook application instance
        outlook = win32com.client.Dispatch("Outlook.Application")

        # Create a new mail item
        mail = outlook.CreateItem(0)  # 0 = olMailItem

        # Set email properties
        mail.To = to
        if cc.strip():
            mail.CC = cc
        mail.Subject = subject
        mail.Body = body

        # Send the email
        mail.Send()

        return True, "Email sent successfully via Outlook"

    except Exception as e:
        return False, f"Error sending email: {str(e)}"
