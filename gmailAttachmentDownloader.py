# Download your attachment files from GMAIL account
# @author - Ratan Lal Prasad
# @email - ratan.kgn@gmail.com

from __future__ import print_function
import httplib2
import os
import platform

import base64
from apiclient import errors

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail Files Downloader'

STORE_DIRECTORY_PATH = '/Users/ratan/Movies/'   # Change it to your local directory path
GMAIL_LABEL_NAME = 'personal-images' # Change it to the label from which you want to download the file


"""Gets valid user credentials from storage.

If nothing has been stored, or if the stored credentials are invalid,
the OAuth2 flow is completed to obtain the new credentials.

Returns:
    Credentials, the obtained credential.
"""

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        # print('Storing credentials to ' + credential_path)
    return credentials

# Download attachments for the given message id
def GetAttachments(service, user_id, msg_id, store_dir):
    # """Get and store attachment from Message with given id.
    #
    # Args:
    #   service: Authorized Gmail API service instance.
    #   user_id: User's email address. The special value "me"
    #   can be used to indicate the authenticated user.
    #   msg_id: ID of Message containing attachment.
    #   store_dir: The directory used to store attachments.
    # """

    if not os.path.exists(store_dir) :
        os.mkdir(store_dir)

    try:
        msg = service.users().messages().get(userId=user_id, id=msg_id).execute()

        for part in msg['payload']['parts']:
            if part['filename']:
                # print('filename is ', part['filename'])
                # print('Part is ', part)

                if 'data' in part['body']:
                    file_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8', 'ignore'))
                else:
                    attachMsg = service.users().messages().attachments().get(userId='me', messageId=msg_id,
                                                                             id=part['body']['attachmentId']).execute()
                    # print('Attach msg is ', attachMsg)
                    file_data = base64.urlsafe_b64decode(attachMsg['data'].encode('utf-8', 'ignore'))

                path = ''.join([store_dir, msg_id+part['filename']])

                try:
                    with open(path, 'wb') as fp:
                         fp.write(file_data)
                         fp.close()
                except:
                    print('Could not store: {file} under {op_sys}.'.format(
                        file=path,
                        op_sys=platform.system()))

    except errors.HttpError:
        print('Error while downloading')


def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    gservice = discovery.build('gmail', 'v1', http=http)

    results = gservice.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'] + ' and ID is --> ' + label['id'])
            if label['name'].upper() == GMAIL_LABEL_NAME.upper():
                print("Name " + label['name'] + "   ID is " + label['id'])
                labelId = label['id']

                messages = gservice.users().messages().list(userId='me', labelIds=[labelId], q='has:attachment').execute();

                # Get Messages with attachments
                for dict in messages['messages']:
                    msgid = dict['id']
                    GetAttachments(service=gservice, user_id='me', msg_id=msgid,
                                   store_dir=STORE_DIRECTORY_PATH+GMAIL_LABEL_NAME+"/")

                break


if __name__ == '__main__':
    main()
