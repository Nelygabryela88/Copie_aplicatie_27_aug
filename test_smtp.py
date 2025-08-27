import smtplib
from email.mime.text import MIMEText

SMTP_HOST = "smtp.vitesco-technologies.net"
SMTP_PORT = 587
SMTP_USER = "svv33684@vitesco.com"
SMTP_PASS = "Schaeffler.2025"
FROM_EMAIL = "calibration.maintenance@vitesco-technologies.net"
TO_EMAIL = "cornelia-gabriela.oprea@vitesco.com"  

msg = MIMEText("Test message from Python!\nIs SMTP working?", "plain", "utf-8")
msg['Subject'] = "SMTP Test from Python"
msg['From'] = FROM_EMAIL
msg['To'] = TO_EMAIL

try:
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(FROM_EMAIL, [TO_EMAIL], msg.as_string())
    print("Test email sent successfully!")
except Exception as e:
    print("Failed to send email:", e)