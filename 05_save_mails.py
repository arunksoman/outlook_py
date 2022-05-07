from outlook_py import Outlook
import os

# absolute path to current folder
download_folder = os.path.abspath('')
msg_file = os.path.join(download_folder, "mail_file.msg")

outlook = Outlook()
ns = outlook.get_namespace()

drafts = outlook.get_default_folder(16)

# Get mail contains attachments
mail_with_attachment = [msg for msg in drafts.Items if msg.Attachments.Count > 0]
for message in mail_with_attachment:
    # Refer https://docs.microsoft.com/en-us/office/vba/api/outlook.olsaveastype for from where 3 came from
    message.SaveAs(msg_file, 3)
    # if you have multiple mails and not want to spoil time to download mails
    break
