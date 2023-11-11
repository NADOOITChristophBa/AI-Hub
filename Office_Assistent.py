import re
import win32com.client as win32
from dotenv import load_dotenv
import os
import openai
import json
import win32com.client
import win32com.client.gencache
from openai import OpenAI

# Load environment variables from .env.
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

# Initialize OpenAI client with your API key
openai.api_key = api_key

# Ensure the generation of COM libraries.
win32.gencache.EnsureDispatch("Outlook.Application")
constants = win32.constants


class Appointment:
    """A custom type for holding appointment details along with email information."""

    def __init__(
        self,
        email=None,
        event_title=None,
        location=None,
        date=None,
        time=None,
        **kwargs,
    ):
        self.email = email
        self.event_title = event_title
        self.location = location
        self.date = date
        self.time = time

        # Assign additional attributes from kwargs
        for key, value in kwargs.items():
            setattr(self, key, value)

    @classmethod
    def from_json(cls, data):
        """Create an Appointment instance from JSON data."""
        try:
            # Parse JSON if it's a string, otherwise use the dictionary directly
            details = json.loads(data) if isinstance(data, str) else data

            # Check and handle invalid data
            if not isinstance(details, dict):
                raise ValueError("Invalid data format")

            return cls(**details)
        except json.JSONDecodeError:
            # Handle JSON decoding error
            return None


class Email:
    """A custom type for holding email details."""

    def __init__(self, subject, body, sender, sender_email, received_time):
        self.subject = subject
        self.body = body
        self.sender = sender
        self.sender_email = sender_email
        self.received_time = received_time

    def __str__(self):
        return f"Email from {self.sender} <{self.sender_email}> received at {self.received_time}: {self.subject}"


def clean_email_content(email_content):
    # Remove URLs from the email content
    email_content = re.sub(r"http\S+", "", email_content)

    # Remove sequences of '<' possibly interspersed with whitespace and newlines
    email_content = re.sub(r"(\s*<\s*)+", " ", email_content)

    # Additional cleanup could go here if needed

    return email_content.strip()


def check_email_contains_appointment(sender_email: Email) -> Appointment | None:
    """Determine if the email is about an appointment and return the details."""

    client = OpenAI()

    # Clean up the email content
    email_content = clean_email_content(sender_email.body)

    # Condensed prompt for the Chat API
    messages = [
        {
            "role": "system",
            "content": "You are a helpful assistant. Return JSON objects in response to queries about appointments.",
        },
        {
            "role": "user",
            "content": "Here is an email subject and content. Determine if it's about an appointment. If so, provide the details in JSON format.",
        },
        {"role": "user", "content": f"Subject: {sender_email.subject}"},
        {"role": "user", "content": f"Content: {email_content}"},
        {
            "role": "user",
            "content": r"Return Format Example: {{'appointment': True, 'details': {'event_title': 'Marketing vs. Werbung - Von den Grundlagen zu deinem Geschäftserfolg', 'attendee_name': 'Christoph Backhaus', 'ticket_type': '1 x Ohne Sitzplatzzuweisung', 'total_amount': 'Kostenlos', 'date': 'Mittwoch, 22. November 2023', 'time': 'von 10:00 Uhr bis 14:00 Uhr (MEZ)', 'location': 'Social Impact Frankfurt', 'order_number': '#8268226239', 'order_date': '8. November 2023', 'contact_email': 'backhauschristoph@gmail.com', 'organizer_contact': 'Kontaktieren Sie den Veranstalter', 'event_platform': 'Eventbrite', 'event_platform_address': '535 Mission Street, 8th Floor, San Francisco, CA 94105'}}}",
        },
    ]

    print(messages)

    response = client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=messages,
        response_format={"type": "json_object"},
        stop=["user:", "system:"],  # Add stops to prevent unnecessary token usage
    )

    print(response)

    # Access the response content
    response_text = response.choices[0].message.content.strip()

    # Parse the response text into a Python dictionary
    try:
        appointment = json.loads(response_text)
        # Check if 'appointment' key is False, if so, return None
        if not appointment.get("appointment"):
            return None
        return Appointment.from_json(data=json.dumps(appointment["details"]))
    except json.JSONDecodeError:
        # If the response is not a valid JSON, return None
        return None


def get_unread_emails_from_outlook(outlook, count=1):
    print("Connecting to Outlook...")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(
        constants.olFolderInbox
    )  # Use the constant for inbox
    print("Getting inbox items...")
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    print("Filtering unread messages...")
    unread_messages = [
        msg for msg in messages if msg.UnRead and msg.Class == constants.olMail
    ]
    print(f"Found {len(unread_messages)} unread mail message(s).")
    emails = []

    for msg in unread_messages[:count]:
        # SenderName gives the display name of the sender
        sender_name = msg.SenderName if hasattr(msg, "SenderName") else "Unknown Sender"
        # SenderEmailAddress gives the actual email address of the sender
        sender_email = (
            msg.SenderEmailAddress
            if hasattr(msg, "SenderEmailAddress")
            else "Unknown Email"
        )
        received_time = (
            msg.ReceivedTime if hasattr(msg, "ReceivedTime") else "Unknown Time"
        )

        print(
            f"Processing email from {sender_name} <{sender_email}> received at {received_time}..."
        )

        email_obj = Email(
            subject=msg.Subject,
            body=msg.Body,
            sender=sender_name,
            sender_email=sender_email,  # Make sure to add this field to your Email class
            received_time=received_time,
        )
        emails.append(email_obj)
        # Be careful with the next line, it will mark your messages as read
        # msg.UnRead = False

    return emails


def send_email_via_outlook(subject, body, recipient):
    """Send an email using the local Outlook instance."""
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.To = recipient
    mail.Send()


def get_read_email_from_unread_email(unread_email: Email):
    appointment = check_email_contains_appointment(unread_email)

    if appointment:
        # Print the appointment details using the new attributes
        print(
            f"Appointment found: {appointment.event_title}, From: {unread_email.sender_email} Location: {appointment.location} Date: {appointment.date} Time:{appointment.time}"
        )
        # Additional details can be printed as needed
    else:
        print(
            f"No appointment in this email: {unread_email.subject}, From: {unread_email.sender}"
        )


if __name__ == "__main__":
    # Example usage

    outlook = win32.Dispatch("Outlook.Application")

    unread_emails = get_unread_emails_from_outlook(
        outlook
    )  # Assuming this function returns a list of Email objects
    for unread_email in unread_emails:
        read_email = get_read_email_from_unread_email(unread_email)
        # Check if the email is about an appointment and get the details

    # Test sending an email
    # subject = "Test Email from AI Hub"
    # body = "This is a test email sent from the AI Hub using local Outlook instance."
    # recipient = "notmymail@outlook.de"
    # send_email_via_outlook(subject, body, recipient)
