import smtplib

SENDER_EMAIL = "varunclg5@gmail.com"
SENDER_PASSWORD = "abcdefghijklmnop"  # put your app password here

with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    print("✅ Login successful!")