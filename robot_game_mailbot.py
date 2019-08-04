import sys
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from base64 import urlsafe_b64encode
import dateutil.parser as parser
import datetime
import time


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify',
'https://www.googleapis.com/auth/spreadsheets.readonly']


def send_message(service, to, subject, message_body='', dryrun=True):
    try:
        message = MIMEText(message_body)
        message['to'] = to
        message['from'] = 'robot-db
@gmail.com'
        message['subject'] = subject
        encoded_message = urlsafe_b64encode(message.as_bytes())
        print(message)
        if not dryrun:
            service.users().messages().send(userId='me', body={'raw': encoded_message.decode()}).execute()
    except Exception as e: 
        print(e)
        pass

def str_clean(string):
    return string.lower().replace(" ", "")

def respond_to_messages(service, bot_db, usr_cred, usr_db, dryrun = True):
    
    # Getting all the unread messages from Inbox
    unread_msgs = service.users().messages().list(userId='me',labelIds=['INBOX', 'UNREAD']).execute()

    if 'messages' in unread_msgs:
        mssg_list = unread_msgs['messages']

        for mssg in mssg_list:
            m_id = mssg['id'] # get id of individual message
            message = service.users().messages().get(userId='me', id=m_id).execute() # fetch the message using API

            payld = message['payload']
            headr = payld['headers']

            for item in headr: # getting the Subject,Time Sent, and Sender
                if item['name'] == 'Subject':
                    msg_subject = str_clean(item['value'])
                elif item['name'] == 'Date':
                    msg_date = item['value']
                    date_parse = (parser.parse(msg_date))
                elif item['name'] == 'From':
                    msg_from = str_clean(item['value'].split('>')[0].split('<')[1])

            print(f'trying {msg_subject} from {msg_from}')
            # make sure message is no longer labeled unread
            service.users().messages().modify(userId='me', id=m_id,body={ 'removeLabelIds': ['UNREAD']}).execute()
            # check for proper credentials
            if msg_from not in usr_cred:
                continue
            print(usr_cred[msg_from])
            
            # check for special "info dump" query
            if msg_subject == 'infodump':
                if usr_cred[msg_from]['role'] != 'bot':
                    continue
                send_message(service,msg_from,f'info','\n'.join([f'{k}: {v}' for k,v in bot_db.items()]),dryrun=dryrun)
                continue

            # check that enough time has elapsed
            if usr_cred[msg_from]['last_query'] and (date_parse - usr_cred[msg_from]['last_query'] < datetime.timedelta(minutes=30)):
                print(f'ignoring request from {msg_from}, not enough time elapsed')
                continue

            # check for proper formatting of all other queries
            if msg_subject.count(':') != 1:
                print(f'ignoring request from {msg_from}, missing colon')
                continue
            task, param = msg_subject.split(':')

            # bot task is only to robotkill
            if usr_cred[msg_from]['role'] == 'bot':
                print('try bot command')
                if 'robot' in task:
                    send_message(service,'gamemaster@gmail.com',f'{msg_from} {task} {param}',dryrun=dryrun)
                    usr_cred[msg_from]['last_query']=date_parse
                else:
                    print(f'ignoring bot request from {msg_from}, tried task {task}:{param}')
            if usr_cred[msg_from]['role'] == 'lead':
                print('try lead command')
                if 'info'!=task or param not in usr_db:
                    print(f'ignoring lead request from {msg_from}, tried task {task}:{param}')
                    continue
                send_message(service,msg_from,f"{msg_subject} = {usr_db[param]}",dryrun=dryrun)
                usr_cred[msg_from]['last_query']=date_parse    
            if usr_cred[msg_from]['role'] == 'eng':
                print('try eng command')
                if task not in bot_db or param not in bot_db[task]:
                    print(f'ignoring eng request from {msg_from}, tried task {task}:{param}')
                    continue
                send_message(service,msg_from,f'{msg_subject} = {bot_db[task][param]}',dryrun=dryrun)
                usr_cred[msg_from]['last_query']=date_parse

    return usr_cred

def get_sheet_data(creds):
    sheets_service = build('sheets', 'v4', credentials=creds)
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId='game_db_spreadsheet_id', range='db!A:I').execute()

    values = result.get('values', [])
    values = [[str_clean(k) for k in v] for v in values]
    features = values[0]
    
    usr_cred = {k[0]:{'role':k[2],'last_query':None} for k in values[1:]}
    usr_db = {k[0]:{**{'title':k[1]},**{features[i]:k[i] for i in range(4,len(k))}} for k in values[1:]}
    bot_db = {k[3]:{features[i]:k[i] for i in range(4,len(k))} for k in values[1:] if k[2]=='bot'}
    return usr_cred, bot_db, usr_db

def main(argv):
    dryrun = True
    if len(argv) > 1 and argv[1] == 'live':
        dryrun=False

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
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
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    usr_cred, bot_db, usr_db = get_sheet_data(creds)

    service = build('gmail', 'v1', credentials=creds)

    while True:
        try:
            usr_cred = respond_to_messages(service,bot_db,usr_cred,usr_db,dryrun)
        except:
            pass
        time.sleep(5)
    

if __name__ == '__main__':
    main(sys.argv)
