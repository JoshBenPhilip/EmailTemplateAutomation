import win32com.client
import os

def email_item(to_address, cc_address, email_body, folder_path, companys, subject):
    ol = win32com.client.Dispatch("Outlook.Application")
    level1Support = ol.Session.Accounts["yourOutlookEmail@yourdomain.com"]


    #size of new email

    ol_mail_item = 0x0

    new_mail = ol.CreateItem(ol_mail_item)

    new_mail._oleobj_.Invoke(*(64209, 0, 8, 0, level1Support))

    new_mail.Subject = subject

    new_mail.To = to_address

    new_mail.CC = cc_address

    new_mail.HTMLBody = email_body

    #Get a list of all files paths in folder
    attachment_paths = []
    folder = os.listdir(folder_path)

    for file_name in folder: 
        file_path = os.path.join(folder_path, file_name)
        attachment_paths.append(file_path)

    for attach_path in attachment_paths:
        new_mail.Attachements.Add(attach_path)

    print("Email attachements have been added")

    new_mail.Display()

    print("email has been displayed")

    #new_mail.Send()

    #print ("email has been sent")

EMAIL_TEMPLATE_HEAD = """<head>
    <title>Email Signature</title>
    <style>
        .signature{
            font-family: Arial, sans-serif;
            line-height: 1.5;
        }
        .signature a {
        text-decoration: none;
        color: #0000EE;
        }
        .confidentiality-notice {
            font-size: 0.9em;
            color: grey;
            margin-top: 9px;
        }
    </style>
</head>
"""

sender_signature = """
    <div class="signature"
        <Regards,<br>
        <strong>Sender Name Example</strong><br>
        Example Job Title<br>

        <p>Example Company Name<br>
        Example Address<br>
        Example Phone<br>
        <a href="http://www.examplecompanysite.com">www.examplecompanysite.com</a></p>

        <p class="confidentiality-notice"> CONFIDENTIALITY NOTICE<br>
        Example confidentiality text.
    </div>
"""


#email templates

EXAMPLE_EMAIL_MESSAGE_A = f"""
<html>
{EMAIL_TEMPLATE_HEAD}
<body>
<p>Hello <b>Example Company</b> Team,</p>
<p>Example email content</p>
{sender_signature}
</body>
</html>
"""


EXAMPLE_EMAIL_MESSAGE_B = f"""
<html>
{EMAIL_TEMPLATE_HEAD}
<body>
<p>Hello <b>Example Company B</b> Team,</p>
<p>Example email content B</p>
{sender_signature}
</body>
</html>
"""

ATTACHMENT_FOLDER_A = os.path.abspath("./attachement_folder_a")
ATTACHMENT_FOLDER_B = os.path.abspath("./attachement_folder_b")

COMPANIES_LIST = [
        {
            "Company A": {
                "to_addresses": "companya@companya.com; companyastaff@companya.com",
                "message": EXAMPLE_EMAIL_MESSAGE_A,
                "attachement_folder": ATTACHMENT_FOLDER_A,
            }
        },
        {
            "Company B": {
                "to_addresses": "companyb@companyb.com; companybstaff@companyb.com",
                "message": EXAMPLE_EMAIL_MESSAGE_B,
                "attachement_folder": ATTACHMENT_FOLDER_B,
            }
        }
]

ticket_num = input("Enter ticket number and then press Enter:   ")



# Loop logic
for company_dict in COMPANIES_LIST:
    company_to_addresses = ""
    company_name_str = ""
    company_message = ""

    for company_name, details in company_dict.items():
        company_to_addresses = details["toAddresses"]
        company_cc_addresses = (
            details["cc_addresses"]
            if "cc_addresses" in details
            else "adam@examplecccompany.com; jan@examplecccompany.com" 
        )
        company_name_str = company_name
        company_message = details["message"]
        company_attachment_folder = details["attachement_folder"]


    ticket_number = ticket_num
    email_subject = (
        f"({company_name_str}) ({ticket_number}) Example email blast"
    )
    email_item(
        company_to_addresses,
        company_cc_addresses,
        company_message,
        company_attachment_folder,
        "example company",
        email_subject,
    )

print("Script has been completed")
