import win32com.client

def get_outlook_mail_metadata():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 represents the Inbox folder

    for mail in inbox.Items:
        sender = mail.SenderName
        receiver = mail.To
        subject = mail.Subject
        sent_time = mail.SentOn

        print(f"Sender: {sender}")
        print(f"Receiver: {receiver}")
        print(f"Subject: {subject}")
        print(f"Sent Time: {sent_time}\n")

get_outlook_mail_metadata()