import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# import pyminizip
import time
from PyPDF2 import PdfReader, PdfWriter
import declarations


FILE_PATH = declarations.SAVED_FILE_PATH
DESTINATION_PATH = declarations.ENCRYPTED_FILE_PATH
SENT_FILE = declarations.SENT_FILES_LIST
PDF_FILE_PASSWORD = declarations.PDF_FILE_PASSWORD
MAIL_SUBJECT = declarations.MAIL_SUBJECT
EMAIL_BODY_TEXT = declarations.EMAIL_BODY_TEXT
EMAIL_SENDER = declarations.EMAIL_SENDER
EMAIL_PASSWORD = declarations.EMAIL_PASSWORD
# If there are multiple reciepients, seperate them by comma
EMAIL_RECEIVER = declarations.EMAIL_RECEIVER
EMAIL_RECEIVER_STRING = ", ".join(EMAIL_RECEIVER) # receivers email ID
print(EMAIL_RECEIVER)
print(EMAIL_RECEIVER_STRING)

file_list = [f for f in os.listdir(path = FILE_PATH)]
sent_path = os.path.join(FILE_PATH, SENT_FILE)
# with open(sent_path, 'a') as f:
# #     f.write('\n' + datetime.now().strftime('%Y%m%d_%H%m'))
#     f.write('\n' + file_list[0])

with open(sent_path, 'r') as f:
    sent_files = f.read().splitlines()
# print(sent_files)

file_unsent_list = list(set(file_list).difference(sent_files))
# print(file_unsent_list)
if file_unsent_list:
    for file_unsent in file_unsent_list:
        print('Sending the file - ', file_unsent)
        file_unsent_path = FILE_PATH + '\\' + file_unsent
        file_unsent_name = file_unsent[:-4]
        file_unsent_format = file_unsent[len(file_unsent_name):]

        # encrypting the file
        reader = PdfReader(file_unsent_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        # adding password
        writer.encrypt(user_password = PDF_FILE_PASSWORD)
        # writing the encrypted file
        destination_file_path = DESTINATION_PATH + '\\' + file_unsent_name + file_unsent_format
        with open(destination_file_path, 'wb') as f:
            writer.write(f)
        # uncomment the below code if we need the encrypted ZIP file but encrypted PDF is advisable.
        # don't forget to change the format into .zip if we want to encrypt as a zipped file
        # pyminizip.compress(file_unsent_path, 'prefix', destination_file_path, ZIP_FILE_PASSWORD, 0)

        time.sleep(3) # sleep for few seconds. Just in case. 

        # Sending Email
        # PRE-REQUISITES - https://towardsdatascience.com/how-to-easily-automate-emails-with-python-8b476045c151
        # Generate Application password in GMAIL. 

        SUBJECT = MAIL_SUBJECT + ' ' + file_unsent_name
        ATTACHMENT = destination_file_path
        email_message = MIMEMultipart()
        email_message['From'] = EMAIL_SENDER
        email_message['To'] = EMAIL_RECEIVER_STRING
        email_message['Subject'] = SUBJECT
        body_part = MIMEText(EMAIL_BODY_TEXT, 'plain')
        email_message.attach(body_part)

        # Attaching the file in the email message
        with open(ATTACHMENT, 'rb') as file:
            email_message.attach(MIMEApplication(file.read(), Name = file_unsent_name + file_unsent_format))

        # sending the email messgae
        session = smtplib.SMTP(host = 'smtp.gmail.com', port = 587)
        session.starttls()
        session.login(user = EMAIL_SENDER, password = EMAIL_PASSWORD)
        text = email_message.as_string()
        session.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, text)
        session.quit()

        time.sleep(1)

        # Update the sent files after sending email
        with open(sent_path, 'a') as f:
            f.write('\n' + file_unsent)
        
        print('file sent !')
else:
    print('no updates')