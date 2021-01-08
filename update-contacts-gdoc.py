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
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
]

# Google object IDs to be used.
template_id = '1EZ87N_6StrbLp_FdG7iriRqXwODcQ3-am1xjKBQk2oM'    # ACATBA Photo Directory TEMPLATE
sheet_id = '1fXHB3I9fryPjsIyptizicTE7B7x1GdMJXxA170zOrJY'       # ACATBA Contact Directory - Dec 2020
pics_dir_id = '1gx5ak18CYgtO6lz9jDEhu3nD3zOIkPP7'               # ACATBA photos for importing
output_id = '1h0qzBSEDkH4Wan-4xKImW6oNCFNVvt7YcdzqMK4Rv7k'      # ACATBA Photo Directory TEMPLATE


def do_cmdline(args, infile_id=None, outfile_id=None):
    """
    Handle all cmdline options.
    """
    # Handle "help" option.
    if "help" in args:
        # Print help info and exit.
        print("No help text yet...")
        exit(0)

    # Define the service objects.
    print("Defining the service object...")
    creds = get_creds(SCOPES)
    doc_service = build('docs', 'v1', credentials=creds)
    sh_service = build('sheets', 'v4', credentials=creds)
    dr_service = build('drive', 'v3', credentials=creds)

    # Handle other options.
    if "update" in args:
        # Update output document with current information.
        update_doc(output_id, doc_svc=doc_service, sht_svc=sh_service)
        data = get_doc(doc_service, output_id)
        print(data)

    if "data" in args:
        # Print spreadsheet data from Google Sheet.
        print("Getting data from spreadsheet...")
        rows = get_sheet(sh_service, sheet_id)
        print(rows)

    if "photos" in args:
        # Print list of photo links gathered from Drive folder.
        pass

    if "template" in args:
        # Print Google Doc template code.
        template = get_doc(doc_service, infile_id)
        pp = pprint.PrettyPrinter(depth=20)
        pp.pprint(template)

    if "body" in args:
        # Print Google Doc body code.
        body = get_doc(doc_service, infile_id)['body']
        pp = pprint.PrettyPrinter(depth=20)
        pp.pprint(body)

    if "outline" in args:
        # Print output outlining the body of the document.
        body = get_doc(doc_service, infile_id)["body"]
        pp = pprint.PrettyPrinter(depth=4)
        pp.pprint(body)

    if "table" in args:
        # Print the table from the template.
        parts = get_doc(doc_service, infile_id)['body']['content']
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
        response = delete_range(doc_service, infile_id, start, end)

    if "delete_row" in args:
        # Delete the selected table row from the template.
        for i, j in enumerate(args):
            if j == 'delete_row':
                del_i = i
        start = args[del_i + 1]
        i_row = args[del_i + 2]
        i_col = args[del_i + 3]
        response = delete_row(doc_service, infile_id, start, i_row, i_col)

    if "abook" in args:
        # Print address book.
        print("Getting data from spreadsheet...")
        rows = get_sheet(sh_service, sheet_id)
        abook = create_abook(rows)
        print(abook)

    if "rows" in args:
        # Print output of 1st table row.
        print("Getting data from spreadsheet...")
        rows = get_sheet(sh_service, sheet_id)
        abook = create_abook(rows)
        rows_out = create_output_rows(abook)
        print(rows_out)

    if "tables" in args:
        # Print output of tables requests.
        print("Getting data from spreadsheet...")
        rows = get_sheet(sh_service, sheet_id)
        abook = create_abook(rows)
        rows_out = create_output_rows(abook)
        requests = row_data(rows_out)
        print(requests)

def update_doc(doc_id, doc_svc=None, sht_svc=None):
    # Get document contents.
    print("Gathering info on existing document...")
    doc_before = get_doc(doc_svc, doc_id)

    # Empty the existing document.
    print("Deleting current contents...")
    end = get_end(doc_before)
    response = delete_all(doc_svc, doc_id, end)

    # Add new content.
    requests = []
    input_rows = get_sheet(sht_svc, sheet_id)
    abook = create_abook(input_rows)
    rows = create_output_rows(abook)
    first_table_index = 1
    last_cell_index = 16
    rows.reverse()
    for row in rows:
        # Insert new doc table row.
        requests.append(row_insert(first_table_index))
        # Modify table row properties.
        #requests.append(row_borders(first_table_index + 1))
        # Insert contact data.
        requests.extend(row_data(row, last_cell_index))
        #break

    # Execute update.
    print("Uploading new content...")
    result = send_requests(doc_svc, doc_id, requests)
    print(result)
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

def get_sheet(svc, sheet_id):
    sheet_range = 'Sheet1!A1:Z500'
    result = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=sheet_range
    ).execute()
    rows = result.get('values', [])
    return rows

def create_abook(rows):
    # Take data output from spreadsheet and build contacts dictionary.
    abook = {}
    for row in rows[1:]:
        # Pad rows if not enough data.
        full_length = 10
        if len(row) < full_length:
            row += [''] * (full_length - len(row))
        # Use data from all columns.
        full_name = f"{row[3]}, {row[4]}"
        abook[full_name] = {}
        for c in range(len(row)):
            abook[full_name][rows[0][c]] = row[c]
    return abook

def get_photos(svc, folder_id):
    pass

def send_requests(svc, doc_id, requests):
    response = svc.documents().batchUpdate(
        documentId=doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def row_insert(index):
    # Insert a "row" (actually a 1x3 table) at the end of the document.
    request = {
        "insertTable": {
            "rows": 2,
            "columns": 3,
            "location": {
                'index': index
            }
        }
    }
    return request

def row_borders(start):
    request = {
        "updateTableCellStyle": {
            "tableCellStyle": {
                "borderLeft": {
                    "color": {
                        "color": {
                            "rgbColor": {
                                "blue": 0.0,
                                "green": 0.0,
                                "red": 0.0,
                            },
                        },
                    },
                    "dashStyle": 'SOLID',
                    "width": {
                        "magnitude": 0.0,
                        "unit": 'PT'
                    },
                },
                "borderRight": {
                    "color": {
                        "color": {
                            "rgbColor": {
                                "blue": 0.0,
                                "green": 0.0,
                                "red": 0.0,
                            },
                        },
                    },
                    "dashStyle": 'SOLID',
                    "width": {
                        "magnitude": 0.0,
                        "unit": 'PT'
                    },
                },
                "borderTop": {
                    "color": {
                        "color": {
                            "rgbColor": {
                                "blue": 0.0,
                                "green": 0.0,
                                "red": 0.0,
                            },
                        },
                    },
                    "dashStyle": 'SOLID',
                    "width": {
                        "magnitude": 0.0,
                        "unit": 'PT'
                    },
                },
                "borderBottom": {
                    "color": {
                        "color": {
                            "rgbColor": {
                                "blue": 0.0,
                                "green": 0.0,
                                "red": 0.0,
                            },
                        },
                    },
                    "dashStyle": 'SOLID',
                    "width": {
                        "magnitude": 0.0,
                        "unit": 'PT'
                    },
                },
            },
            "fields": '*',
            "tableStartLocation": {
                "index": start,
            }
        }
    }
    return request

def create_output_rows(abook):
    # Add rows as necessary, populate contact data into each row.
    teams = {}
    for p in abook:
        try:
            teams[abook[p]['Team']].append(p)
        except KeyError:
            teams[abook[p]['Team']] = [p]

    # Create list of 3-item rows with all relevant info.
    rows = []
    sections = []
    for team, members in teams.items():
        # Handle Admin first.
        if team == "Admin":
            sections.insert(0, {team: members})
        elif team == "Finance":
            sections.insert(1, {team: members})
        else:
            sections.append({team: members})

    for section in sections:
        for team, members in section.items():
            sec_rows = []
            sec_row_ct = len(members) // 3 + 1
            for i, member in enumerate(members):
                ind = i // 3
                try:
                    sec_rows[ind].append(member)
                except IndexError:
                    sec_rows.insert(ind, [member])
            rows.append(sec_rows)

    output_rows = []
    for section in rows:
        for row in section:
            # Create a new row to send to output.
            output_rows.append(row)

    final_rows = []
    for names_row in output_rows:
        abook_row = []
        for n in names_row:
            abook_row.append(abook[n])
        final_rows.append(abook_row)

    return final_rows

def row_data(row, index):
    requests = []
    #print(row)
    qty = len(row)
    if qty < 3:
        pass
    for i in range(qty):
        # Adjust factor "f" so that entries are inserted at the beginning of the
        #   row rather than the end.
        f = i + 3 - qty
        full_name = f"{row[i]['Name 1']}, {row[i]['Name 2']}"
        team = row[i]['Team']
        title = row[i]['Role']
        emails = f"{row[i]['Email 1']}, {row[i]['Email 2']}"
        skype = row[i]['Skype Name']
        phones = f"{row[i]['Phone 1']}, {row[i]['Phone 2']}"
        text = f"{full_name}\n{team}, {title}\nEmail: {emails}\nSkype: {skype}\nTel: {phones}"
        requests.append({
            'insertText': {
                "text": text,
                "location": {
                    "index": index - 2 * f
                }
            }
        })
    print(requests)
    return requests

def populate_row(svc, doc_id):
    # Add content to an existing empty table of 2 rows and 3 columns.
    pass

def delete_all(svc, doc_id, end):
    response = None
    if end > 1:
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
