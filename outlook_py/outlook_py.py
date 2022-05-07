import win32com.client as winclient

class Outlook:
    def __init__(self):
        self.outlook = winclient.Dispatch('outlook.application')
        self.is_open = True

    def create_item(self, itemtype):
        """
        Go through createitem documentation  https://docs.microsoft.com/en-us/office/vba/api/outlook.application.createitem
        """
        # You can find out itemtype here: https://docs.microsoft.com/en-us/office/vba/api/outlook.olitemtype
        self.create_item = self.outlook.CreateItem(itemtype)

    def get_namespace(self):
        """
        https://docs.microsoft.com/en-us/office/vba/api/outlook.application.getnamespace
        """
        self.outlook_namespace = self.outlook.GetNameSpace('MAPI')
        return self.outlook_namespace

    def get_default_folder(self, folder_enumeration):
        return self.outlook_namespace.GetDefaultFolder(folder_enumeration)

    def send_mail(self, mail_options, body_format=0, sensitivity=0):
        """
        Go throught mail item and properties
        """
        self.create_item.To = mail_options['to']
        self.create_item.Subject = mail_options['subject']
        if mail_options.get('html_body', None):
            self.create_item.HTMLBody = mail_options['html_body']
        self.create_item.Body = mail_options['body']
        for f in mail_options.get("attachments", []):
            self.create_item.Attachments.Add(f)
        # check this link https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat
        self.create_item.BodyFormat = body_format
        # Chec this link https://docs.microsoft.com/en-us/office/vba/api/outlook.olsensitivity
        self.create_item.Sensitivity = sensitivity
        self.create_item.Send()

    def msg_file_handle(self, namespace, file_path):
        return namespace.OpenSharedItem(file_path)


    def quit(self):
        if self.is_open:
            self.outlook.Quit()
