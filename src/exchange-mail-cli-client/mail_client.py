import os
import msal
import requests
import json
from datetime import datetime
from dotenv import load_dotenv

from html_converter import HTMLToTextConverter

load_dotenv()

# Конфигурация
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.ReadWrite.Shared",
]
CACHE_FILE = "token_cache.bin"
MAILBOX = "me"

# Инициализация кэша токенов
cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_FILE):
    cache.deserialize(open(CACHE_FILE, "r").read())

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)


def save_cache():
    """Сохраняет кэш токенов"""
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def acquire_token():
    """Получает токен доступа"""
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            return result

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError("Failed to create device flow")

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    save_cache()
    return result


def send_mail(
    access_token, subject, body, to_recipients, cc_recipients=None, save_to_sent=True
):
    """Отправляет электронное письмо"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/sendMail"

    email_msg = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [
                {"emailAddress": {"address": address}} for address in to_recipients
            ],
        },
        "saveToSentItems": save_to_sent,
    }

    # Добавляем копии если есть
    if cc_recipients:
        email_msg["message"]["ccRecipients"] = [
            {"emailAddress": {"address": address}} for address in cc_recipients
        ]

    headers = {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/json",
    }

    try:
        response = requests.post(endpoint, json=email_msg, headers=headers)
        if response.status_code == 202:
            print("✅ Email sent successfully!")
            return True
        else:
            print(f"❌ Failed to send email: {response.status_code} - {response.text}")
            return False
    except Exception as e:
        print(f"❌ Error sending email: {str(e)}")
        return False


def get_emails(access_token, top=10, folder="inbox"):
    """Получает список писем из указанной папки"""
    endpoint = (
        f"https://graph.microsoft.com/v1.0/{MAILBOX}/mailFolders/{folder}/messages"
    )
    params = {
        "$top": top,
        "$orderby": "receivedDateTime DESC",
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments",
    }

    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.get(endpoint, headers=headers, params=params)
        if response.status_code == 200:
            emails = response.json().get("value", [])
            print(f"\n📥 Found {len(emails)} emails in {folder}:")

            for i, email in enumerate(emails, 1):
                read_status = "📖" if email.get("isRead", False) else "📨"
                attachment_status = "📎" if email.get("hasAttachments", False) else ""
                from_info = email.get("from", {}).get("emailAddress", {})
                from_address = from_info.get("address", "Unknown")
                from_name = from_info.get("name", from_address)
                subject = email.get("subject", "No subject")
                date = email.get("receivedDateTime", "")[:19].replace("T", " ")

                print(f"{i:2d}. {read_status}{attachment_status} {subject}")
                print(f"     From: {from_name} | Date: {date} | ID: {email['id']}")

            return emails
        else:
            print(f"❌ Failed to get emails: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        print(f"❌ Error getting emails: {str(e)}")
        return []


def get_email_content(access_token, message_id):
    """Получает полное содержимое письма"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/messages/{message_id}"
    params = {
        "$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,body,receivedDateTime,bodyPreview,hasAttachments,importance"
    }

    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.get(endpoint, headers=headers, params=params)
        if response.status_code == 200:
            email_data = response.json()
            return process_email_content(email_data)
        else:
            print(
                f"❌ Failed to get email content: {response.status_code} - {response.text}"
            )
            return None
    except Exception as e:
        print(f"❌ Error getting email content: {str(e)}")
        return None


def process_email_content(email_data):
    """Обрабатывает данные письма для отображения"""
    body = email_data.get("body", {})
    content_type = body.get("contentType", "text")
    content = body.get("content", "")

    # Конвертируем HTML в читаемый текст
    if content_type == "html":
        readable_content = HTMLToTextConverter.convert(content)
    else:
        readable_content = content

    # Извлекаем информацию о вложениях
    attachments_info = []
    if content_type == "html":
        attachments_info = HTMLToTextConverter.extract_attachments_info(content)

    return {
        "id": email_data.get("id"),
        "subject": email_data.get("subject", "No subject"),
        "from": email_data.get("from", {})
        .get("emailAddress", {})
        .get("address", "Unknown"),
        "from_name": email_data.get("from", {}).get("emailAddress", {}).get("name", ""),
        "to_recipients": [
            recipient.get("emailAddress", {}).get("address", "")
            for recipient in email_data.get("toRecipients", [])
        ],
        "cc_recipients": [
            recipient.get("emailAddress", {}).get("address", "")
            for recipient in email_data.get("ccRecipients", [])
        ],
        "bcc_recipients": [
            recipient.get("emailAddress", {}).get("address", "")
            for recipient in email_data.get("bccRecipients", [])
        ],
        "received_date": email_data.get("receivedDateTime"),
        "content_type": content_type,
        "readable_content": readable_content,
        "body_preview": email_data.get("bodyPreview", ""),
        "has_attachments": email_data.get("hasAttachments", False),
        "attachments_info": attachments_info,
        "importance": email_data.get("importance", "normal"),
    }


def display_email_content(email_content):
    """Отображает содержимое письма в читаемом формате"""
    if not email_content:
        print("❌ No email content to display")
        return

    # Заголовок
    print("\n" + "=" * 80)
    importance_symbol = (
        "🔴"
        if email_content["importance"] == "high"
        else "🟡" if email_content["importance"] == "low" else "🔵"
    )
    print(f"{importance_symbol} SUBJECT: {email_content['subject']}")
    print("=" * 80)

    # Информация об отправителе и получателях
    print(f"📧 FROM: {email_content['from_name']} <{email_content['from']}>")
    print(f"📨 TO: {', '.join(email_content['to_recipients'])}")

    if email_content["cc_recipients"]:
        print(f"📋 CC: {', '.join(email_content['cc_recipients'])}")

    if email_content["bcc_recipients"]:
        print(f"📋 BCC: {len(email_content['bcc_recipients'])} recipients")

    print(f"📅 DATE: {email_content['received_date']}")

    # Информация о вложениях
    attachment_status = "✅ Yes" if email_content["has_attachments"] else "❌ No"
    print(f"📎 ATTACHMENTS: {attachment_status}")

    if email_content["attachments_info"]:
        print(f"📋 MENTIONED ATTACHMENTS: {len(email_content['attachments_info'])}")
        for att in email_content["attachments_info"]:
            print(f"   - {att['name']}")

    print("-" * 80)

    # Preview если есть
    if email_content["body_preview"]:
        print(f"📝 PREVIEW: {email_content['body_preview']}")
        print("-" * 80)

    # Основное содержимое
    print("📄 CONTENT:")
    print("-" * 80)
    print(email_content["readable_content"])
    print("=" * 80)


def delete_email(access_token, message_id):
    """Удаляет письмо"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/messages/{message_id}"

    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.delete(endpoint, headers=headers)
        if response.status_code == 204:
            print("✅ Email deleted successfully!")
            return True
        else:
            print(
                f"❌ Failed to delete email: {response.status_code} - {response.text}"
            )
            return False
    except Exception as e:
        print(f"❌ Error deleting email: {str(e)}")
        return False


def move_email_to_trash(access_token, message_id):
    """Перемещает письмо в корзину"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/messages/{message_id}/move"

    headers = {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/json",
    }

    data = {"destinationId": "deleteditems"}

    try:
        response = requests.post(endpoint, headers=headers, json=data)
        if response.status_code == 201:
            print("✅ Email moved to trash successfully!")
            return True
        else:
            print(
                f"❌ Failed to move email to trash: {response.status_code} - {response.text}"
            )
            return False
    except Exception as e:
        print(f"❌ Error moving email to trash: {str(e)}")
        return False


def search_emails(access_token, query, top=10):
    """Ищет письма по запросу"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/messages"
    params = {
        "$top": top,
        "$search": f'"{query}"',
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments",
    }

    headers = {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/json",
    }

    try:
        response = requests.get(endpoint, headers=headers, params=params)
        if response.status_code == 200:
            emails = response.json().get("value", [])
            print(f"\n🔍 Found {len(emails)} emails for query '{query}':")

            for i, email in enumerate(emails, 1):
                read_status = "📖" if email.get("isRead", False) else "📨"
                attachment_status = "📎" if email.get("hasAttachments", False) else ""
                from_info = email.get("from", {}).get("emailAddress", {})
                from_address = from_info.get("address", "Unknown")
                subject = email.get("subject", "No subject")

                print(f"{i:2d}. {read_status}{attachment_status} {subject}")
                print(f"     From: {from_address} | ID: {email['id']}")

            return emails
        else:
            print(f"❌ Search failed: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        print(f"❌ Error searching emails: {str(e)}")
        return []


def get_folders(access_token):
    """Получает список папок почтового ящика"""
    endpoint = f"https://graph.microsoft.com/v1.0/{MAILBOX}/mailFolders"

    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            folders = response.json().get("value", [])
            print("\n📁 Available folders:")
            for folder in folders:
                print(f"  - {folder['displayName']} (ID: {folder['id']})")
            return folders
        else:
            print(f"❌ Failed to get folders: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        print(f"❌ Error getting folders: {str(e)}")
        return []


def main():
    """Главная функция приложения"""
    print("🚀 Microsoft Graph Mail Client")
    print("Initializing...")

    # Получаем токен
    result = acquire_token()

    if "access_token" not in result:
        print(
            "❌ Authentication failed:",
            result.get("error"),
            result.get("error_description"),
        )
        return

    access_token = result["access_token"]
    print("✅ Authentication successful!")

    # Главный цикл
    while True:
        print("\n" + "=" * 50)
        print("📧 MICROSOFT GRAPH MAIL MANAGER")
        print("=" * 50)
        print("1. 📤 Send email")
        print("2. 📥 Read inbox")
        print("3. 👀 Read email content")
        print("4. 🗑️ Delete email")
        print("5. 🗂️ Move email to trash")
        print("6. 🔍 Search emails")
        print("7. 📁 List folders")
        print("8. 🚪 Exit")

        choice = input("\nSelect option (1-8): ").strip()

        if choice == "1":
            # Отправка письма
            subject = input("Enter subject: ").strip()
            body = input("Enter message: ").strip()
            to_emails = input("Enter recipient emails (comma separated): ").split(",")
            to_emails = [email.strip() for email in to_emails if email.strip()]

            cc_emails = input("Enter CC emails (comma separated, optional): ").split(
                ","
            )
            cc_emails = [email.strip() for email in cc_emails if email.strip()]

            if not to_emails:
                print("❌ At least one recipient is required")
                continue

            send_mail(
                access_token, subject, body, to_emails, cc_emails if cc_emails else None
            )

        elif choice == "2":
            # Чтение inbox
            limit = input("Number of emails to show (default 10): ").strip()
            try:
                limit = int(limit) if limit else 10
            except ValueError:
                limit = 10

            emails = get_emails(access_token, top=limit)

        elif choice == "3":
            # Чтение содержимого письма
            message_id = input("Enter message ID: ").strip()
            if message_id:
                email_content = get_email_content(access_token, message_id)
                if email_content:
                    display_email_content(email_content)
            else:
                print("❌ Message ID is required")

        elif choice == "4":
            # Полное удаление письма
            message_id = input("Enter message ID to delete: ").strip()
            if message_id:
                confirm = (
                    input("⚠️ Are you sure? This cannot be undone! (y/n): ")
                    .strip()
                    .lower()
                )
                if confirm == "y":
                    delete_email(access_token, message_id)
            else:
                print("❌ Message ID is required")

        elif choice == "5":
            # Перемещение в корзину
            message_id = input("Enter message ID to move to trash: ").strip()
            if message_id:
                move_email_to_trash(access_token, message_id)
            else:
                print("❌ Message ID is required")

        elif choice == "6":
            # Поиск писем
            query = input("Enter search query: ").strip()
            if query:
                search_emails(access_token, query)
            else:
                print("❌ Search query is required")

        elif choice == "7":
            # Список папок
            get_folders(access_token)

        elif choice == "8":
            print("👋 Goodbye!")
            break

        else:
            print("❌ Invalid option. Please try again.")


if __name__ == "__main__":
    main()
