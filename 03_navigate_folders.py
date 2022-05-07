from outlook_py import Outlook

outlook = Outlook()
ns = outlook.get_namespace()

# https://docs.microsoft.com/en-us/office/vba/api/outlook.folder

# Default folders enumeration: Like inbox, sent, draft etc. 
# https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

inbox = outlook.get_default_folder(6)
print(inbox)
for folder in inbox.Folders:
    print(folder)

account = ns.Folders['arunkavilkarottus55@outlook.com']
for folder in account.Folders:
    print(folder)
