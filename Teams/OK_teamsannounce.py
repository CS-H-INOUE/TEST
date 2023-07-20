# Memo
#
# 簡単な通知は可能
# メンションやファイルのアップロードはGraph　-> API使用
#

import requests

def send_teams_notification(webhook_url, message):
    headers = {'Content-Type': 'application/json'}
    data = {'text': message}

    response = requests.post(webhook_url, json=data, headers=headers)

    if response.status_code == 200:
        print("Notification sent successfully.")
    else:
        print(f"Failed to send notification. Status code: {response.status_code}, Response: {response.text}")

if __name__ == "__main__":
    # TeamsのWebhook URLを設定してください
    teams_webhook_url = "https://chubusunsho.webhook.office.com/webhookb2/ac7b5ae5-fd1c-4264-a1fa-6b1d7bb7b38e@e60fb0e2-e4db-4ea3-a60f-888216fb15cd/IncomingWebhook/5b3fb26cfe0b4ac0a07ef6c041330892/d5dbd0cb-710e-44bd-8d17-4b3be9317246"
    message = "Hello from Python! This is a test notification."

    send_teams_notification(teams_webhook_url, message)
