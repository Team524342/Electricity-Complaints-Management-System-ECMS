import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import speech_recognition as sr

def send_email_smtp(sender_email, receiver_email, subject, message, password, smtp_server="smtp.gmail.com", smtp_port=587, html_message=None):
    """
    Send an email using SMTP protocol.

    Parameters:
    - sender_email: Sender's email address
    - receiver_email: Recipient's email address
    - subject: Email subject
    - message: Email body content (plain text)
    - html_message: Email body content (HTML)
    - password: Sender's email password or app password
    - smtp_server: SMTP server address (default: Gmail's SMTP server)
    - smtp_port: SMTP server port (default: 587 for TLS)

    Returns:
    - Boolean indicating success/failure
    """
    try:
        # Setup the MIME
        email_message = MIMEMultipart("alternative")
        email_message['From'] = sender_email
        email_message['To'] = receiver_email
        email_message['Subject'] = subject

        # Attach the plain text and HTML message to the email
        email_message.attach(MIMEText(message, 'plain'))
        if html_message:
            email_message.attach(MIMEText(html_message, 'html'))

        # Create a secure connection with the server and send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()  # Identify ourselves to the SMTP server
            server.starttls()  # Secure the connection
            server.ehlo()  # Re-identify after starting TLS
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, email_message.as_string())

        print(f"Email sent successfully to {receiver_email}")
        return True

    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

# # Example usage
# if __name__ == "__main__":
#     sender = "mohammadferozali7866@gmail.com"
#     # receiver = "maferoz7866@gmail.com"
#     subject = "Electricity Complaint Status Update"

#     complaint_id = "EC123456"
#     customer_name = "John Doe"
#     complaint_date = "2024-05-10"
#     status = "Resolved"
#     resolution_time = "2024-05-11 14:30"
#     support_contact = "1800-123-456"

#     message = (
#         f"Dear {customer_name},\n\n"
#         f"Thank you for contacting the Electricity Board.\n"
#         f"Your complaint (ID: {complaint_id}) registered on {complaint_date} has been updated.\n\n"
#         f"Current Status: {status}\n"
#         f"Resolution Time: {resolution_time}\n\n"
#         f"If you have further issues, please contact our support at {support_contact}.\n\n"
#         f"Thank you,\n"
#         f"Electricity Board Support Team"
#     )

#     html_message = f"""
#     <html>
#     <head>
#       <style>
#         body {{
#           font-family: Arial, sans-serif;
#           background-color: #f7f7f7;
#           margin: 0;
#           padding: 0;
#         }}
#         .container {{
#           background: #fff;
#           max-width: 600px;
#           margin: 40px auto;
#           padding: 30px 40px;
#           border-radius: 8px;
#           box-shadow: 0 2px 8px rgba(0,0,0,0.08);
#         }}
#         h2 {{
#           color: #2d7be5;
#         }}
#         .details {{
#           background: #f0f4fa;
#           padding: 15px 20px;
#           border-radius: 6px;
#           margin: 20px 0;
#         }}
#         .footer {{
#           font-size: 13px;
#           color: #888;
#           margin-top: 30px;
#         }}
#       </style>
#     </head>
#     <body>
#       <div class="container">
#         <h2>Electricity Complaint Status Update</h2>
#         <p>Dear <strong>{customer_name}</strong>,</p>
#         <p>Thank you for contacting the Electricity Board.<br>
#         Your complaint details are as follows:</p>
#         <div class="details">
#           <p><strong>Complaint ID:</strong> {complaint_id}<br>
#           <strong>Date Registered:</strong> {complaint_date}<br>
#           <strong>Status:</strong> {status}<br>
#           <strong>Resolution Time:</strong> {resolution_time}</p>
#         </div>
#         <p>If you have further issues, please contact our support at <strong>{support_contact}</strong>.</p>
#         <div class="footer">
#           Thank you,<br>
#           Electricity Board Support Team
#         </div>
#       </div>
#     </body>
#     </html>
#     """

#     password = "gtdglxjdyinudaxy"  # For Gmail, use an App Password

#     send_email_smtp(sender, receiver, subject, message, password, html_message=html_message)

#     # Speech recognition example
#     r = sr.Recognizer()
#     with sr.AudioFile('yourfile.wav') as source:
#         audio = r.record(source)
