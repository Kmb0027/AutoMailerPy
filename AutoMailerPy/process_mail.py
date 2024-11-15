import win32com.client
import time

# Initialize Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def send_email(recipient, subject, body):
    mail = outlook.CreateItem(0)  # 0 is the code for MailItem
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    mail.Send()
    print(f"Email sent to {recipient} with subject '{subject}'.")

def check_inbox(sender_email, subject_filter=None):
    inbox = namespace.GetDefaultFolder(6)  # 6 is the code for the inbox
    messages = inbox.Items
    messages = messages.Restrict("[SenderEmailAddress] = '{}'".format(sender_email))
    if subject_filter:
        messages = messages.Restrict("[Subject] = '{}'".format(subject_filter))

    for message in messages:
        print(f"From: {message.SenderEmailAddress}")
        print(f"Subject: {message.Subject}")
        print(f"Body: {message.Body[:100]}...")  # print the first 100 characters of the body
        print("===")
        # Mark message as read
        message.UnRead = False

# Example usage
recipient = "someone@example.com"
subject = "Automated Email"
body = "Hello, this is an automated email sent via Python."

# Send an email
send_email(recipient, subject, body)

# Wait a few seconds, then check the inbox
time.sleep(5)

# Check for new emails from a specific sender
check_inbox("sender@example.com", subject_filter="Automated Email")