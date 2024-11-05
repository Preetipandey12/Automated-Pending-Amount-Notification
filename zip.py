import xlrd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Path to Excel file
file_path = r'C:\Users\ppree\OneDrive\Documents\Preeti.xlsx'  # Replace with the actual path to your Excel file
workbook = xlrd.open_workbook(file_path)

# Access the specific sheet by name
sheet = workbook.sheet_by_name("PreetiPandey")  # Replace with your actual sheet name

# Calculate the total pending amount
total_pending_amount = 450  # Initialize with a starting amount if needed
for row_idx in range(1, sheet.nrows):  # Start from 1 to skip header row
    row = sheet.row(row_idx)  # Get the entire row
    pending_amount = row[3].value  # Assume the pending amount is in the 4th column (index 3)
    if isinstance(pending_amount, (int, float)) and pending_amount > 0:
        total_pending_amount += pending_amount

# Email configuration
smtp_server = 'smtp.gmail.com'
port = 587
sender_email = 'ppppreety029@gmail.com'  # Replace with your email
password = 'ofyn sfvl vdhd qnws'         # Replace with your app-specific password if you have 2FA enabled
receiver_email = 'ppreety029@gmail.com'  # Replace with the recipient's email

# Email content
subject = "Pending Amount Summary"
body = f"Dear Customer,\n\nThe total pending amount as per the records is: {total_pending_amount}.\nPlease take the necessary steps to clear the pending amount.\n\nThank you."

# Create the email message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

# Send the email
try:
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()  # Start TLS for security
        server.login(sender_email, password)  # Log in to your email account
        server.send_message(msg)  # Send the email
    print("Email sent successfully!")
except Exception as e:
    print(f"Error occurred while sending email: {e}")
