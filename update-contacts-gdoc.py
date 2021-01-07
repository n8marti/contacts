#!/usr/bin/env python3

import os.path
import pickle
import sys
import pprint

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/documents',
]

# Google Docs IDs to be used.
template_id = '1EZ87N_6StrbLp_FdG7iriRqXwODcQ3-am1xjKBQk2oM'
output_id = '1h0qzBSEDkH4Wan-4xKImW6oNCFNVvt7YcdzqMK4Rv7k'


def parse_cmdline(args, infile_id=None, outfile_id=None):
    """
    Handle all cmdline options.
    """
    # Define the service object.
    creds = get_creds(SCOPES)
    service = build('docs', 'v1', credentials=creds)

    # Handle options.
    if "data" in args:
        # Print spreadsheet data from Google Sheet.
        pass

    if "photos" in args:
        # Print list of photo links gathered from Drive folder.
        pass

    if "template" in args:
        # Print Google Doc template code.
        template = get_doc(service, infile_id)
        pp = pprint.PrettyPrinter(depth=20)
        pp.pprint(template)

    if "cell" in args:
        # Print output describing a single table cell
        pass

    if "outline" in args:
        # Print output outlining the body of the document.
        body = get_doc(service, infile_id)["body"]
        pp = pprint.PrettyPrinter(depth=4)
        pp.pprint(body)

def get_creds(scopes):
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
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_doc(svc, doc_id):
    # Retrieve the documents contents from the Docs service.
    doc_dict = svc.documents().get(documentId=doc_id).execute()
    return doc_dict

def get_sheet(svc, doc_id):
    pass

def get_photos(svc, folder_id):
    pass

def insert_table(svc, doc_id, index):
    request1 = {
        "insertTable": {
            "rows": 10,
            "columns": 1,
            "location": {
                "index": index,
            }
        }
    }

    requests = [request1]
    response = svc.documents().batchUpdate(
        documentId=doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def delete_range(svc, doc_id, start, end):
    request1 = {
        "deleteContentRange": {
            "range": {
                "startIndex": start,
                "endIndex": end,
            }
        }
    }
    requests = [request1]
    response = svc.documents().batchUpdate(
        documentId=doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def main():
    """
    Creates a Google Doc file that merges togehter photos and contact info
    gleaned from other shared Drive files.
    """
    parse_cmdline(sys.argv, infile_id=template_id, outfile_id=output_id)

    #insert_table(service, DOCUMENT_ID, 2)
    #for start in outline.keys():
    #    if outline[start]["type"] == "table":
    #        end = outline[start]["end"]
    #        delete_range(service, DOCUMENT_ID, start, end)
    #get_outline(doc)


if __name__ == '__main__':
    main()
