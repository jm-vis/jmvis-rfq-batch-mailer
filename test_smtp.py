import os, ssl, smtplib
from dotenv import load_dotenv

load_dotenv()  # liest .env im Projekt

host = os.getenv("SMTP_HOST")
port = int(os.getenv("SMTP_PORT", "587"))
user = os.getenv("SMTP_USER")
pw   = os.getenv("SMTP_PASSWORD")
use_ssl = os.getenv("USE_SSL", "false").lower() == "true"

print(f"Testing SMTP login: host={host} port={port} ssl={use_ssl} user={user}")

ctx = ssl.create_default_context()
try:
    if use_ssl:
        s = smtplib.SMTP_SSL(host, port, context=ctx, timeout=30)
    else:
        s = smtplib.SMTP(host, port, timeout=30)
        s.ehlo()
        s.starttls(context=ctx)
        s.ehlo()
    code, resp = s.login(user, pw)
    print("LOGIN OK:", code, resp)
    s.quit()
except smtplib.SMTPAuthenticationError as e:
    print("AUTH FAIL:", e.smtp_code, e.smtp_error)
except Exception as e:
    print("OTHER ERROR:", repr(e))
