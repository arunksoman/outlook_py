from outlook_py import Outlook
import os

outlook = Outlook()
ns = outlook.get_namespace()

drafts = outlook.get_default_folder(16)

# Get mail contains attachments
mail_with_attachment = [msg for msg in drafts.Items if msg.Attachments.Count > 0]

for msg in mail_with_attachment:
    attachments = msg.Attachments
    for attachment in attachments:
        # print(dir(attachment))
        """
        attachment has these properties
        ['AddRef', 'Application', 'BlockLevel', 'Class', 'Delete', 'DisplayName', 'FileName', 'GetIDsOfNames', 'GetTemporaryFilePath', 'GetTypeInfo', 'GetTypeInfoCount', 'Index', 'Invoke', 'MAPIOBJECT', 'Parent', 'PathName', 'Position', 'PropertyAccessor', 'QueryInterface', 'Release', 'SaveAsFile', 'Session', 'Size', 'Type']
        """
        # print(attachment.FileName)
        attachment.SaveAsFile(os.path.join(os.path.abspath(''), attachment.FileName))
    break