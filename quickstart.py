from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import pandas as pd
import re
import base64
import csv
import datetime

time = datetime.datetime.now()
time = time.strftime("%b_%d_%H")

# #### Next Step:
# - ver casos em que fica .xls
# - arrajar forma de gravar


# More information in https://developers.google.com/gmail/api/v1/reference/

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

path_token = r'../token_email/token.pickle'
path_csv = r'C:\JFCImportantes\Knowledge\Engineering\Automation\hr_dbase\hr_'
path_csv = path_csv + time + '.csv'


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    path_token = r'../token_email/token.pickle'
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(path_token):
        with open(path_token, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(r'../token_email/token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    # emails = service.users().messages().get()

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            if label['name'] == "MI_HR":
                print("MI_HR_exists")
                labelid = label['id']
                print(labelid)
                break
        # ------------- Get the data in the messages
        gmail_query = "subject: exportNotify label:MI_HR label:inbox"
        # The subject is specified by the app, and -label:inbox excludes
        # archived emails
        messages = service.users().messages().list(userId='me',
                                                   q=gmail_query).execute()

        if messages['resultSizeEstimate'] > 1:  # if multiple messages exist

            # last message
            id = messages["messages"][messages['resultSizeEstimate']-1]["id"]
            print(id)

            # Block to retrieve message
            message = service.users().messages().get(userId='me', id=id)
            message = message.execute()

            # Block to retrieve attachment
            for part in message['payload']['parts']:
                if part['filename']:
                    if 'data' in part['body']:
                        data = part['body']['data']
                    else:
                        att_id = part['body']['attachmentId']
                        att = service.users().messages().attachments()
                        att = att.get(userId='me', messageId=id, id=att_id)
                        att = att.execute()
                        data = att['data']
                        # print(data)

                    # Retrieve data from attachment
                    ext = re.search('exportNotify.(.*)',
                                    message["payload"]["headers"][3]['value'])
                    print(ext[1])

                    if ext[1] == "csv":
                        str_csv = base64.urlsafe_b64decode(data.encode('UTF-8'))
                        str_csv = str_csv.decode("utf-8")
                        # print(str_csv)
                        with open(path_csv, "w") as csv_file:
                            # Create the writer object with tab delimiter
                            writer = csv.writer(csv_file, delimiter=';')
                            writer.writerow(str_csv)

                    elif ext[1] == "xls":
                        data_xls = pd.read_excel(data)
                        print("niiiice")
                    else:
                        print("not xls or csv")

            # Confirm the order of the messages
            # dict = {'ids': [], 'dates': []}
            # timestamp = 0
            # for message in messages["messages"]:
            #     id = message["id"]
            #     print(id)
            #     dict['ids'].append(message["id"])
            #     msg = service.users().messages().get(userId='me', id=id,
            #                                          format='metadata').execute()
            #
            #     dict['ids'].append(msg["internalDate"])
            #
            #     if timestamp < int(msg["internalDate"]):
            #         last_id = id
            # print(re.search(',(.*) -080', msg['payload']['headers'][1]['value']).group(0))
            # print(last_id)


        elif messages['resultSizeEstimate'] > 0: # If just one message exists
            print(messages["messages"][0]["id"]["partid"])



if __name__ == '__main__':
    main()
