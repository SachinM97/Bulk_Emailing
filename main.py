import os
import click
import pandas as pd
import time
import tkinter
from tkinter import filedialog
import docx2txt
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import ntpath


def get_files_path(sender_email_address, sender_email_password):
    # get current working directory
    try:
        cwd = os.getcwd()
        input_dir = tkinter.Tk()
        input_dir.withdraw()
        # pick up the Input file path
        print("Please select the input files directory")
        print("---------------------------------------\n")
        time.sleep(1)
        tempdir = filedialog.askdirectory(parent=input_dir, initialdir=cwd, title='Please select a directory')
        if len(tempdir) > 0:
            if '\\' in tempdir:
                tempdir = tempdir + '\\'
            else:
                tempdir = tempdir + '/'
            print("You choose:\n", tempdir)
        list_of_docx = []
        count = 1
        print("\nThe files present are:\n")
        for root, dirs, files in os.walk(tempdir):
            for file in files:
                if file.endswith(".xlsx"):
                    print(f"{count}. {file}")
                    list_of_docx.append(file)
                    count = count + 1
        file_number_input_file = count + 1
        while int(file_number_input_file) >= count:
            file_number_input_file = input("\nPlease type the file number for Input File: ")
            if int(file_number_input_file) >= count:
                print("\nSorry you have inputted wrong value. Please type the correct value again.")
        filename_input_file = list_of_docx[int(file_number_input_file) - 1]
        input_file_path = tempdir + filename_input_file
        if ".xlsx" not in input_file_path:
            input_file_path = input_file_path + ".xlsx"
        # pick up the email draft path
        list_of_docx = []
        count = 1
        print("\nThe files present are:\n")
        for root, dirs, files in os.walk(tempdir):
            for file in files:
                if file.endswith(".docx"):
                    print(f"{count}. {file}")
                    list_of_docx.append(file)
                    count = count + 1
        file_number_email = count + 1
        while int(file_number_email) >= count:
            file_number_email = input("\nPlease type the file number for Email Doc: ")
            if int(file_number_email) >= count:
                print("\nSorry you have inputted wrong value. Please type the correct value again.")
        filename_email = list_of_docx[int(file_number_email) - 1]
        input_email_path = tempdir + filename_email
        if ".docx" not in input_email_path:
            input_email_path = input_email_path + ".docx"
        email_template = docx2txt.process(input_email_path)
        # pick up the attachment directory
        print("\nPlease select the attachment directory")
        print("--------------------------------------\n")
        time.sleep(1)
        tempdir = filedialog.askdirectory(parent=input_dir, initialdir=cwd, title='Please select a directory')
        if len(tempdir) > 0:
            if '\\' in tempdir:
                tempdir = tempdir + '\\'
            else:
                tempdir = tempdir + '/'
            print("You choose:\n", tempdir)
        list_of_docx = []
        for root, dirs, files in os.walk(tempdir):
            for file in files:
                if file.endswith(".xlsx"):
                    list_of_docx.append(file)
        read_excel(sender_email_address, sender_email_password, input_file_path, tempdir, list_of_docx, email_template)
    except Exception as ErrorInPickingFiles:
        print("The files are not imported properly.\nError in importing input files : {}".format(ErrorInPickingFiles))


# read the input file
def read_excel(sender_email_address, sender_email_password, input_file_path, attachment_path, list_of_docx, email_template):
    try:
        valid_supplier_cols = range(0, 3)
        df = pd.read_excel(input_file_path, sheet_name=0, header=0, usecols=valid_supplier_cols)
        i = 0
        for row in df['Recipient email address']:
            to_recipient = row.lower()
            email_subject = df['Email Subject'][i]
            if any(row in s for s in list_of_docx):
                attachment = attachment_path + row
                if '.xlsx' not in attachment:
                    attachment = attachment + '.xlsx'
            else:
                attachment = None
            cc_recipients = df['CC recipient'][i]
            cc_recipients = cc_recipients.split(',')
            cc_recipients = [x.strip(' ') for x in cc_recipients]
            send_emails(sender_email_address, sender_email_password, email_subject, to_recipient, cc_recipients, email_template, attachment)
            print("Processed {} files".format(i+1))
            i = i+1
    except Exception as ErrorInReadingFiles:
        print("The input files format is not correct. Please provide the correct format and retry.\nError in reading input files : {}".format(ErrorInReadingFiles))


def send_emails(sender_email_address, sender_email_password, email_subject, to_recipient, cc_recipients, email_template, attachment):
    try:
        if len(to_recipient) > 0:
            msg = MIMEMultipart()
            msg['From'] = sender_email_address
            msg['To'] = to_recipient
            msg['Subject'] = email_subject
            if len(cc_recipients) > 0:
                to_addrs = [to_recipient] + cc_recipients
            else:
                to_addrs = [to_recipient]
            print(f"Sending Emails to: {to_addrs}")
            msg['Cc'] = ";".join(cc_recipients)

            msg.attach(MIMEText(email_template, 'html'))
            head, tail = ntpath.split(attachment)
            attachment_send = open(attachment, "rb")
            p = MIMEBase('application', "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            p.set_payload(attachment_send.read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', 'attachment', filename=tail)
            msg.attach(p)
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login(sender_email_address, sender_email_password)
            text = msg.as_string()
            s.sendmail(sender_email_address, to_addrs, text)
            s.quit()
        else:
            print("\nNo recipient found for this row")
    except Exception as ErrorInSendingEmails:
        print("Error in sending out email : {}".format(ErrorInSendingEmails))


@click.command()
@click.option('--sender_email_address', prompt='Sender Email Address', help='Email address you want to send the supplier recon emails from')
@click.option('--sender_email_password', prompt='Sender Email Password', help='Password for the same email address provided')
def get_inputs_from_user(sender_email_address, sender_email_password):
    try:
        if sender_email_address is not None and sender_email_password is not None:
            get_files_path(sender_email_address, sender_email_password)
    except Exception as CredentialsError:
        print("Sender's email id and password not provided.\nError in receiving Sender's data : {}".format(CredentialsError))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    get_inputs_from_user()