from outlook_py import Outlook
import os

# absolute path to current folder
download_folder = os.path.abspath('')
msg_file = os.path.join(download_folder, "mail_file.msg")

outlook = Outlook()
ns = outlook.get_namespace()
my_gmail = ns.Folders['arunksoman5678@gmail.com']

for folder in my_gmail.Folders:
    print(folder)
