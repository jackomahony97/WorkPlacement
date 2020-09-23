from io import StringIO
from html.parser import HTMLParser
import imaplib
import email
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import re


month_to_pass = ""


def connectToSheets(row):
    """
    Get day, month year
    """

    """
    Connect to google sheet
    """
    # use creds to create a client to interact with the Google Drive API
    try:
        toTest = int(orderNum)
        if toTest >= 10300 :
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
                      "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', SCOPES)
            client = gspread.authorize(creds)

            # Find a workbook by name and open the first sheet
            # Make sure you use the right name here.

            sheet = client.open("Customers").worksheet("Email List")

            try:
                sheet.find(orderNum)

            except:
                sheet.append_row(row)
    except:
        print("failed")


EMAIL = ""
PASSWORD = ""
SERVER = ""

class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()

    def handle_data(self, d):
        self.text.write(d)

    def get_data(self):
        return self.text.getvalue()


def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()


while 1 == 1:
    try:
        datetimeInfo = datetime.datetime.now()

        year = datetimeInfo.year

        month = datetimeInfo.month
        if month < 10:
            month = "0" + str(month)
        else:
            month = str(month)

        day = datetimeInfo.day

        if day < 10:
            day = "0" + str(day)
        else:
            day = str(day)

        sheetTitle = str(day) + "/" + month + "/" + str(year)
        # connect to the server and go to its inbox
        mail = imaplib.IMAP4_SSL(SERVER)
        mail.login(EMAIL, PASSWORD)
        # we choose the inbox but you can select others

        mail.select('inbox')
        # we'll search using the ALL criteria to retrieve
        # every message inside the inbox
        # it will return with its status and a list of ids
        status, data = mail.search(None, '(SINCE "03-Jun-2020")')
        # the list returned is a list of bytes separated
        # by white spaces on this format: [b'1 2 3', b'4 5 6']
        # so, to separate it first we create an empty list
        mail_ids = []
        # then we go through the list splitting its blocks
        # of bytes and appending to the mail_ids list
        for block in data:
            # the split function called without parameter
            # transforms the text or bytes into a list using
            # as separator the white spaces:
            # b'1 2 3'.split() => [b'1', b'2', b'3']
            mail_ids += block.split()

        # now for every id we'll fetch the email
        # to extract its content
        for mail_id in mail_ids:
            # the fetch function fetch the email given its id
            # and format that you want the message to be
            status, data = mail.fetch(mail_id, '(RFC822)')

            # the content data at the '(RFC822)' format comes on
            # a list with a tuple with header, content, and the closing
            # byte b')'
            for response_part in data:
                # so if its a tuple...
                if isinstance(response_part, tuple):
                    # we go for the content at its second element
                    # skipping the header at the first and the closing
                    # at the third
                    message = email.message_from_bytes(response_part[1])

                    # with the content we can extract the info about
                    # who sent the message and its subject
                    mail_from = message['from']
                    mail_subject = message['subject']
                    mail_date = email.utils.parsedate_tz(message['date'])
                    mail_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(mail_date))
                    month_to_pass = str(mail_date.month)

                    mail_date = str(mail_date.day) + "/" + str(mail_date.month) + "/" + str(mail_date.year)

                    # then for the text we have a little more work to do
                    # because it can be in plain text or multipart
                    # if its not plain text we need to separate the message
                    # from its annexes to get the text
                    if message.is_multipart():
                        mail_content = ''

                        # on multipart we have the text message and
                        # another things like annex, and html version
                        # of the message, in that case we loop through
                        # the email payload
                        for part in message.get_payload():
                            # if the content type is text/plain
                            # we extract it
                            if part.get_content_type() == 'text/plain':
                                if mail_from == '"PersonaliseMe@TheBabyShopCork" <info@personaliseme.ie>':
                                    mail_content += part.get_payload()
                                elif mail_from == "The Embroidery Hut <info@theembroideryhut.com>":
                                    mail_content += str(part.get_payload(decode=True).decode('utf-8'))

                    else:
                        mail_content = ""
                        # if the message isn't multipart, just extract it
                        if mail_from == '"PersonaliseMe@TheBabyShopCork" <info@personaliseme.ie>':
                            mail_content = message.get_payload()
                        elif mail_from == '"The Embroidery Hut @ The Baby Shop, Cork" <info@theembroideryhut.com>':
                            mail_content = message.get_payload()
                        elif mail_from == '"Personalise Me @ The Baby Shop, Cork" <info@personaliseme.ie>':
                            mail_content = message.get_payload()

                    pass_email = "email not found"
                    if mail_content != "":



                        mail_content = strip_tags(mail_content)
                        body = mail_content
                        mail_content = ' '.join(mail_content.split())
                        tosesh = mail_content
                        mail_content = mail_content.split()

                        if re.search(r'\bcustomer order\b', tosesh):

                            orderNum = re.compile( "\((.*)\)" ).search( mail_subject ).group( 1 )



                            i = 0
                            done = False

                            billingAddress = ""

                            for i in range(len(mail_content) - 1):
                                if mail_content[i] == "Billing" and done == False:
                                    j = i + 2
                                    while mail_content[j] != "Shipping" and j < len(mail_content) - 1:
                                        billingAddress += mail_content[j] + " "
                                        j += 1
                                    done = True
                                i += 1

                            pass_email = re.findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", str(mail_content))

                            if len(pass_email) == 0:
                                pass_email = "not found"

                            connectToSheets([orderNum, mail_date, pass_email[0]])

                            # and then let's show its result
                            print(f'From: {mail_from}')
                            print(f'OrderNum: {orderNum}')

                            print(billingAddress)
                            print(f'Email: {pass_email}')
                            print(f'Subject: {mail_subject}')
                            print(f'Date: {mail_date}')
                            print(f'Content: {mail_content}\n')


        mail.close()
        mail.logout()
    except:
        print("failed 2")
        continue

