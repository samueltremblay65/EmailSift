"""This module is used to send test emails to an inbox."""

import smtplib, ssl

port = 465  # For SSL
password = "zmll grlo jqyl wkee"

# Create a secure SSL context
context = ssl.create_default_context()

def send_email(email):
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login("emailsift2006@gmail.com", password)
        
        sender_email = "emailsift2006@gmail.com"
        receiver_email = email.get("receiver")
        content = email.get("content")
        subject = email.get("subject")

        message = "Subject: {0}\n\n{1}".format(subject, content)

        server.sendmail(sender_email, receiver_email, message)


email = {"subject": "Check out these deals", "content": "This should be the main content of the email", "receiver": "samuel.tremblay65@gmail.com"}

def send_email_from_commandline():
    receiver = input("Enter the name of the recipient: ")
    subject = input("Subject line: ")
    content = input("Message: ")

    email = {"receiver": receiver, "subject": subject, "content": content}

    send_email(email)
