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


def do_cmdline(args, infile_id=None, outfile_id=None):
    """
    Handle all cmdline options.
    """
    # Handle "help" option.
    if "help" in args:
        # Print help info and exit.
        print("No help text yet...")
        exit(0)

    # Define the service object.
    print("Defining the service object...")
    creds = get_creds(SCOPES)
    service = build('docs', 'v1', credentials=creds)

    # Handle other options.
    if "update" in args:
        # Update output document with current information.
        update_doc(service, output_id)

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

    if "body" in args:
        # Print Google Doc body code.
        body = get_doc(service, infile_id)['body']
        pp = pprint.PrettyPrinter(depth=20)
        pp.pprint(body)

    if "cell" in args:
        # Print output describing a single table cell
        pass

    if "outline" in args:
        # Print output outlining the body of the document.
        body = get_doc(service, infile_id)["body"]
        pp = pprint.PrettyPrinter(depth=4)
        pp.pprint(body)

    if "row" in args:
        # Print one row of output from the body.
        pass

    if "table" in args:
        # Print the table from the template.
        parts = get_doc(service, infile_id)['body']['content']
        for part in parts:
            try:
                pp = pprint.PrettyPrinter(depth=20)
                pp.pprint(part['table'])
            except KeyError:
                pass

    if "delete_range" in args:
        # Delete the selected range from the template.
        for i, j in enumerate(args):
            if j == 'delete_range':
                del_i = i
        start = args[del_i + 1]
        end = args[del_i + 2]
        response = delete_range(service, infile_id, start, end)

    if "delete_row" in args:
        # Delete the selected table row from the template.
        for i, j in enumerate(args):
            if j == 'delete_row':
                del_i = i
        start = args[del_i + 1]
        i_row = args[del_i + 2]
        i_col = args[del_i + 3]
        response = delete_row(service, infile_id, start, i_row, i_col)

def update_doc(svc, doc_id):
    # Get document contents.
    print("Gathering info on existing document...")
    doc_before = get_doc(svc, doc_id)

    # Empty the existing document.
    print("Deleting current contents...")
    end = get_end(doc_before)
    response = delete_all(svc, doc_id, end)

    # Add new content.
    # TODO: Need to build up a list of requests to apply all at once.
    print("Adding new content...")
    response = row_insert(svc, doc_id)
    #row_borders(svc, doc_id)

    print("Done.")

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

def get_end(doc):
    # Return index of end position.
    last = doc['body']['content'][-1]['endIndex']
    # Have to back up one from "last" to preserve trailing newline.
    end = last - 1
    return end

def get_sheet(svc, doc_id):
    pass

def get_photos(svc, folder_id):
    pass

def row_insert(svc, doc_id):
    # Insert a "row" (actually a 1x3 table) at the end of the document.
    requests = [
        {
            "insertTable": {
                "rows": 2,
                "columns": 3,
                "endOfSegmentLocation": {
                    "segmentId": ''
                }
            }
        }
    ]
    response = svc.documents().batchUpdate(
        documentId=doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def row_borders(svc, doc_id):
    requests = [
        {
            "tableCellStyle": {
                "tableStartLocation": {
                    blahblah,
                }
            }
        }
    ]
    response = svc.documents().batchUpdate(
        documentId = doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def populate_row(svc, doc_id):
    # Add content to an existing empty table of 2 rows and 3 columns.
    pass

def delete_all(svc, doc_id, end):
    requests = [
        {
            "deleteContentRange": {
                "range": {
                    "startIndex": 1,
                    "endIndex": end,
                }
            }
        }
    ]
    response = svc.documents().batchUpdate(
        documentId = doc_id,
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
        documentId = doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def delete_row(svc, doc_id, start, i_row, i_col):
    requests = [
        {
            'deleteTableRow': {
                'tableCellLocation': {
                    'tableStartLocation': {
                        'index': start
                    },
                    'rowIndex': i_row,
                    'columnIndex': i_col
                }
            }
        }
    ]
    response = svc.documents().batchUpdate(
        documentId = doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def main():
    """
    Creates a Google Doc file that merges togehter photos and contact info
    gleaned from other shared Drive files.
    """
    do_cmdline(sys.argv, infile_id=template_id, outfile_id=output_id)


if __name__ == '__main__':
    main()
