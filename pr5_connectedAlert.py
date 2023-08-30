import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# Read data from Excel file into a DataFrame
excel_file_path = 'pr5_Internship_data.xlsx'
df = pd.read_excel(excel_file_path)

# Filter for poor performance
poor_performance_df = df[df['Performance'] == 'Poor']

# Send an email if there are poor performances
if not poor_performance_df.empty:
    # Email configuration (replace with your actual credentials)
    sender_email = 'SENDER_EMAIL'
    sender_password = 'PASSWORD'
    recipient_email = 'RECIPIENT_EMAIL'
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Create email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = 'Poor Performance Report'

    # Create HTML table from DataFrame
    table_html = poor_performance_df.to_html(index=False)

    # Attach the HTML table to the email
    msg.attach(MIMEText(table_html, 'html'))

    try:
        # Connect to SMTP server and send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print("Error sending email:", str(e))
else:
    print("No poor performances to report.")
