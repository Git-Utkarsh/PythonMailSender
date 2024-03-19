import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import getpass

def login_outlook_email(email, password):
    # Outlook email credentials
    outlook_email = email
    outlook_password = password

    # Outlook SMTP server details
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587  # Outlook SMTP port

    try:
        # Establishing connection with Outlook server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Start TLS encryption
        # Login to Outlook server
        server.login(outlook_email, outlook_password)
        print("Login successful!")
        return server
    except Exception as e:
        print("Login failed. Error:", str(e))
        return None

def send_email(server, sender_email, receiver_email, subject, body_html, attachment_filename=None):
    try:
        # Create a multipart message
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject

        # Attach HTML body
        message.attach(MIMEText(body_html, 'html'))

        if attachment_filename:
            # Attach file
            with open(attachment_filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {attachment_filename}",
            )
            message.attach(part)

        # Send the message
        server.sendmail(sender_email, receiver_email, message.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print("Email sending failed. Error:", str(e))

# Example usage:
if __name__ == "__main__":
    # Input Outlook email credentials
    email = input("Enter your Outlook email: ")
    password = getpass.getpass("Enter your Outlook password: ")

    # Log in to Outlook email
    server = login_outlook_email(email, password)
    if server:
        body_html = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GitHub Payload</title>
<style>
    .container {
        border: 2px solid black;
        padding: 20px;
        background-color: #f0f0f0;
        width: 300px;
    }
    h2 {
        text-align: center;
    }
    .button {
        display: block;
        width: 100%;
        background-color: #007bff;
        color: white;
        text-align: center;
        padding: 10px 0;
        text-decoration: none;
        border-radius: 5px;
    }
</style>
</head>
<body>
<div class="container">
    <h2>Get Google Search </h2>
    <p style="text-align: center;">Click the button below to view</p>
    <a href="https://google.com" class="button">View Google(30 words)</a>
</div>
</body>
</html>
"""

        # Send email
        sender_email = email
        receiver_email = input("Enter recipient's email address: ")
        subject = input("Enter email subject: ")
        attachment_filename = input("Enter attachment filename (leave empty if none): ").strip()
        send_email(server, sender_email, receiver_email, subject, body_html, attachment_filename)
        server.quit()  # Close the connection
