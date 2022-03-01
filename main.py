# Python code to send email to a list of
# emails from a spreadsheet

import argparse
# load env variables
import os
import re
import sys
import smtplib
from time import sleep
from email.mime.text import MIMEText
from email.header import Header

# import the required libraries
import xlrd
from dotenv import load_dotenv
from tqdm import trange

load_dotenv(dotenv_path='config.txt')


# Detect if we take the file from exist directory or from binary file (in EXE case)
def resource_path(relative_path):
    """
    :param relative_path: string, path the software want
    :return: string, return the path with the binary or regular path
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def slowPrint(text):
    if __name__ == "__main__":
        for letter in text:
            print(letter, end='')
            sleep(3. / 90)
    else:
        print(text)


def sendEmail(login_email, login_password, spreadsheet, email_subject, email_body, from_email, direction):
    slowPrint("Welcome to the email sender")
    print("\r")
    slowPrint("This program will send emails to a list of emails from a excel sheet")
    print("\n")

    slowPrint("Trying to login\r")
    # establishing connection with gmail
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(login_email, login_password)

    slowPrint("Login successful")
    print("\r")

    names = []
    emails = []

    # reading the spreadsheet
    workbook = xlrd.open_workbook(spreadsheet, encoding_override='utf8')
    worksheet = workbook.sheet_by_index(0)
    data = []
    first_row = ["EMAIL"]  # Header
    for col in range(1, worksheet.ncols):
        first_row.append(worksheet.cell_value(0, col))

    for row in range(1, worksheet.nrows):
        elm = {}
        for col in range(worksheet.ncols):
            if first_row[col] == "EMAIL":
                emails.append(worksheet.cell_value(row, col))
            else:
                elm[first_row[col]] = worksheet.cell_value(row, col)
        data.append(elm)

    if emails is None:
        raise ValueError("EMAIL not exist in the excel file")

    slowPrint("Reading the excel file is successful")
    print("\r")

    slowPrint("Reading the message file is successful")
    print("\n")

    slowPrint("Please ensure that the message and the excel is correct")
    print("\r")
    slowPrint("Press enter to continue, else, exit the program => ")
    input()
    slowPrint("Start to send the emails\n")

    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    # iterate through the records
    for i in trange(len(emails)):
        # for every record get the name and the email addresses
        email = emails[i]

        if not re.fullmatch(regex, email):
            continue

        if direction == "rtl":
            messageText = f'<html dir="rtl" lang="he">' \
                          f'<p style="white-space: pre-wrap;text-align:right;direction: rtl;">' \
                          f'{email_body}' \
                          f'</p>' \
                          f'</html>'
        else:
            messageText = f'<html>' \
                          f'<p style="white-space: pre-wrap;">' \
                          f'{email_body}' \
                          f'</p>' \
                          f'</html>'

        for key, value in data[i].items():
            # If value is not string, then convert it to string and delete all that after .
            if not isinstance(value, str):
                value = str(value)
                value = value.split('.')[0]

            # replace the name and email in the message
            messageText = messageText.replace(f'{{{key}}}', value)

        # Set the message to the file contents
        message = MIMEText(messageText, 'html', 'utf-8')
        message['Subject'] = Header(email_subject, 'utf-8')
        message['From'] = Header(from_email, 'utf-8')

        # sending the email
        server.sendmail(login_email, [email], message.as_string())

    # close the smtp server
    server.close()

    slowPrint("All emails are sent")
    print("\r")


def emailer(email=None, password=None, spreadsheet=None, subject=None, body=None, from_email=None, direction=None):
    parser = argparse.ArgumentParser(description='Send emails from excel file')
    parser.add_argument('-e', '--email', type=str, help='Email address')
    parser.add_argument('-p', '--password', type=str, help='Password')
    parser.add_argument('-s', '--spreadsheet', type=str, help='Spreadsheet file')
    parser.add_argument('-t', '--subject', type=str, help='Email subject')
    parser.add_argument('-b', '--body', type=str, help='Email body or path to the file')
    parser.add_argument('-f', '--from_email', type=str, help='From email')
    parser.add_argument('-d', '--dir', type=str, help='Direction of the text')

    args = parser.parse_args()

    # If email not exist, look for config.txt
    if args.email is None and email is None:
        if not os.path.isfile("config.txt"):
            if getattr(sys, 'frozen', False):
                os.system(f'copy "{resource_path("help/config.txt")}" config.txt')
                os.system(f'copy "{resource_path("help/readme.md")}" readme.txt')
                slowPrint("Please read the readme file and configure the config.txt file")
                sys.exit()
            else:
                raise ValueError("No value was provided for email and config.txt not exist")

    email = email or args.email or os.getenv("EMAIL")
    if email is None or email == "":
        raise ValueError("No value was provided for email")

    password = password or args.password or os.getenv("EMAIL_PASSWORD")
    if password is None or password == "":
        raise ValueError("No value was provided for password")

    spreadsheet = spreadsheet or args.spreadsheet or os.getenv("EXCEL_FILE_NAME")
    if spreadsheet is None or subject == "":
        raise ValueError("No value was provided for spreadsheet")
    elif not os.path.isfile(spreadsheet):
        spreadsheet = f'./{spreadsheet}.xlsx'
        if not os.path.isfile(spreadsheet):
            raise ValueError("Spreadsheet file not exist")

    subject = subject or args.subject or os.getenv("SUBJECT")
    if subject is None or subject == "":
        subject = "Email"
        print("No value was provided for subject, use default value")

    body = body or args.body or os.getenv("MESSAGE_FILE_NAME")
    if body is None or body == "":
        raise ValueError("No value was provided for body")

    # Check if body is file
    if os.path.isfile(body):
        with open(body, 'r', encoding='utf8') as f:
            body = f.read()
    elif os.path.isfile(f'./{body}.txt'):
        with open(f'./{body}.txt', 'r', encoding='utf8') as f:
            body = f.read()

    from_email = from_email or args.from_email or os.getenv("FROM")
    if from_email is None or from_email == "":
        from_email = email

    direction = direction or args.dir or os.getenv("DIRECTION")
    if direction is None or direction == "":
        direction = "ltr"

    sendEmail(email, password, spreadsheet, subject, body, from_email, direction)


if __name__ == "__main__":
    emailer()
