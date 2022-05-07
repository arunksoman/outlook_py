from outlook_py import Outlook
import os

# absolute path to current folder
download_folder = os.path.abspath('')
msg_file = os.path.join(download_folder, "mail_file.msg")

outlook = Outlook()
ns = outlook.get_namespace()
file_path = msg_file

msg = outlook.msg_file_handle(ns, file_path)

print(f"Sender Name: {msg.SenderName}")
print(f"Sender Mail: {msg.SenderEmailAddress}")
print(f"Send On: {msg.SentOn}")
print(f"To: {msg.To}")
print(f"CC: {msg.CC}")
print(f"Subject: {msg.Subject}")
print(f"Body: {msg.Body}")
