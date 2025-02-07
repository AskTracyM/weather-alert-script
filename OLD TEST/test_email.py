import smtplib
from email.mime.text import MIMEText

# Test email credentials
EMAIL_ADDRESS = "notifymbfs@gmail.com"  # Replace with your email
EMAIL_PASSWORD = "tvkj aiwg nbzq kalo"           # Replace with your password
RECIPIENT_EMAIL = "admin@mortgagebankersfs.com"  # Replace with your recipient's email

try:
    # Create email content
    msg = MIMEText("This is a test email.")
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = RECIPIENT_EMAIL
    msg["Subject"] = "Test Email"

    # Send email
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, msg.as_string())
    print("Test email sent successfully!")
except Exception as e:
    print(f"Error sending test email: {e}")
