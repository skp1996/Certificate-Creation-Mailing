import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
form = pd.read_excel('feedback_form.xlsx')


# The mail addresses and  password
sender_address ="sk27808@gmail.com"
sender_pass ="7789004313"

# list of reciver email_id to the mail
li = form['Email'].to_list()
name = form['Name'].to_list()
# [item for item in input("Enter Receiver Mail Address :- ").split()] this is used to take user input of receiver mail id


# getting length of list
length = len(li)

# Iterating the index
# same as 'for i in range(len(list))'

# Here we iterate the loop and send msg one by one to the reciver
for i in range(length):
    X = li[i]
    reciver_mail = X

    print(reciver_mail)

    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = reciver_mail  # Pass Reciver Mail Address
    message['Subject'] = 'Mail using python'  # The subject line

    mail_content = '''Hello %s,
    This is a mail from Sourav Kumar Purohit. There is a certificate of participation file attachments in the mail.
    Thank You''' %(name[i])

    # The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'plain'))
    # Create SMTP session for sending the mail
    # open the file to be sent
    filename ="certificate_"+str(name[i])+".pdf";
    # Open PDF file in binary mode
    # The file is in the directory same as where you run your Python script code from
    with open(filename, "rb") as attachment:
        # MIME attachment is a binary file for that content type "application/octet-stream" is used
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # encode into base64
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'part' to instance 'message'
    message.attach(part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(sender_address, sender_pass)
    text = message.as_string()
    s.sendmail(sender_address, reciver_mail, text)
    s.quit()

    print('Mail Sent')

# It Send Separated Mail one by one each receiver mail