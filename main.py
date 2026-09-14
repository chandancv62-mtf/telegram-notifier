import os
import requests

BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_TO")

message = "🎉 Hello! Yeh mera pehla message hai jo GitHub Actions aur Python se aaya hai!"

url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
payload = {"chat_id": CHAT_ID, "text": message}

response = requests.post(url, data=payload)
print("Status:", response.status_code)
