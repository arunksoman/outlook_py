from outlook_py import Outlook

outlook = Outlook()
# item type is 0 for mail object
outlook.create_item(0)

mail_data = {
    "to": "arunksoman5678@gmail.com",
    "subject": "Test mail",
    "html_body": "<h1>Outlook Test</h1>",
    "body": "hmm.. Hello this mail is generated using python code"
}
outlook.send_mail(mail_options=mail_data, body_format=2, sensitivity=2)
print("[Info] Message sent successfully")
# outlook.quit()
