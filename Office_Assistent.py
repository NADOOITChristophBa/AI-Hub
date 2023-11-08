import win32com.client as win32


def send_email_via_outlook(subject, body, recipient):
    # Start an instance of Outlook
    outlook = win32.Dispatch("outlook.application")

    # Create a new email
    mail = outlook.CreateItem(0)

    # Set the email details
    mail.Subject = subject
    mail.Body = body
    mail.To = recipient

    # Send the email
    mail.Send()


def get_unread_emails():
    # Start an instance of Outlook
    outlook = win32.Dispatch("outlook.application")
    namespace = outlook.GetNamespace("MAPI")

    # Access the Inbox
    inbox = namespace.GetDefaultFolder(6)  # 6 is the folder number for Inbox
    messages = inbox.Items

    # Filter for unread emails
    unread_messages = [msg for msg in messages if msg.UnRead]

    # Print subject and sender of each unread email
    for msg in unread_messages:
        print(f"Subject: {msg.Subject}, From: {msg.SenderName}")


# Replace with your subject, email body, and recipient email
subject = "Test Email from AI Hub"
body = "This is a test email sent from the AI Hub using local Outlook instance."
recipient = "notmymail@outlook.de"  # Replace with your email address

# Send the email
send_email_via_outlook(subject, body, recipient)

# Get and print the list of unread emails
get_unread_emails()
