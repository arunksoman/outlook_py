from outlook_py import Outlook

outlook = Outlook()
ns = outlook.get_namespace()

# Create/ Add new folder
inbox = outlook.get_default_folder(6)
inbox.Folders.Add('test1')
inbox.Folders.Add('test2')


# Deleting Folder
inbox.Folders['test2'].Delete()
