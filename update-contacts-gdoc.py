#!/usr/bin/env python3

import pickle
import sys
import pprint

from pathlib import Path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
]

### Google object IDs to be used.

# ACATBA Photo Directory TEMPLATE
# https://docs.google.com/document/d/1EZ87N_6StrbLp_FdG7iriRqXwODcQ3-am1xjKBQk2oM
template_id = '1EZ87N_6StrbLp_FdG7iriRqXwODcQ3-am1xjKBQk2oM'

# ACATBA Contact Directory - Dec 2020
# https://docs.google.com/spreadsheets/d/1fXHB3I9fryPjsIyptizicTE7B7x1GdMJXxA170zOrJY
sheet_id = '1fXHB3I9fryPjsIyptizicTE7B7x1GdMJXxA170zOrJY'

# ACATBA photos for importing
# https://drive.google.com/drive/folders/1gx5ak18CYgtO6lz9jDEhu3nD3zOIkPP7
pics_dir_id = '1gx5ak18CYgtO6lz9jDEhu3nD3zOIkPP7'

# ACATBA Photo Directory
# https://docs.google.com/document/d/1h0qzBSEDkH4Wan-4xKImW6oNCFNVvt7YcdzqMK4Rv7k
output_id = '1h0qzBSEDkH4Wan-4xKImW6oNCFNVvt7YcdzqMK4Rv7k'


def do_cmdline(args, infile_id=None, outfile_id=None):
    """Handle all cmdline options."""

    # Handle "help" option.
    for help in {'help', '--help', '-h'}:
        if help in args:
            # Print help info and exit.
            print("No help text yet...")
            exit(0)

    # Handle options.
    # "Update" is default option if none are given.
    if "update" in args or len(args) == 1:
        """Update output document with current information."""

        # Build services.
        svc_dict = build_services(['docs', 'sheets', 'drive'])
        doc_svc = svc_dict['docs']
        sh_svc = svc_dict['sheets']
        dr_svc = svc_dict['drive']

        # Update document.
        update_doc(outfile_id, doc_svc=doc_svc, sht_svc=sh_svc, dir_svc=dr_svc)
        exit()

    if "data" in args:
        """Print file content from doc, sheet, or template."""
        file_type = 'sheets' # default option
        obj_id = sheet_id
        if len(args) > 2:
            for i, arg in enumerate(args):
                if arg == "data":
                    object = args[i-1]
                    if object in {'doc', 'docs', 'template'}:
                        file_type = 'docs'
                        if object == 'template':
                            obj_id = infile_id
                        else:
                            obj_id = outfile_id
                    elif object in {'sheet', 'sheets'}:
                        file_type = 'sheets'
                        obj_id = sheet_id
                    elif object in {'drive', 'photo', 'photos'}:
                        file_type = 'drive'
                        obj_id = pics_dir_id
                    else:
                        print(f"Bad file type given ({object}).")
                        exit(1)
                    break

        # Build services.
        svc_dict = build_services([file_type])
        svc = svc_dict[file_type]

        # Print data from file_type.
        print(f"Getting data from {object}...")
        if file_type == 'sheets':
            data = get_sheet(svc, obj_id)
        elif file_type == 'docs':
            data = get_doc(svc, obj_id)
        elif file_type == 'drive':
            data = get_photos(svc, obj_id)
        print(data)
        exit()

    if "body" in args:
        """Print Google Doc body code."""
        # Build services.
        svc_dict = build_services(['docs'])
        svc = svc_dict['docs']
        # Get content.
        body = get_doc(svc, infile_id)['body']
        pp = pprint.PrettyPrinter(depth=20)
        pp.pprint(body)

    if "outline" in args:
        """Print output outlining the body of the document."""
        # Build services.
        svc_dict = build_services(['docs'])
        svc = svc_dict['docs']
        # Get content.
        body = get_doc(svc, infile_id)["body"]
        pp = pprint.PrettyPrinter(depth=4)
        pp.pprint(body)

    if "table" in args:
        """Print the table from the template."""
        # Build services.
        svc_dict = build_services(['docs'])
        svc = svc_dict['docs']

        parts = get_doc(svc, infile_id)['body']['content']
        for part in parts:
            try:
                pp = pprint.PrettyPrinter(depth=20)
                pp.pprint(part['table'])
            except KeyError:
                pass

    if "delete_range" in args:
        """Delete the selected range from the template."""
        # Build services.
        svc_dict = build_services(['docs'])
        svc = svc_dict['docs']

        for i, j in enumerate(args):
            if j == 'delete_range':
                del_i = i
        start = args[del_i + 1]
        end = args[del_i + 2]
        response = delete_range(svc, infile_id, start, end)

    if "delete_row" in args:
        """Delete the selected table row from the template."""
        # Build services.
        svc_dict = build_services(['docs'])
        svc = svc_dict['docs']

        for i, j in enumerate(args):
            if j == 'delete_row':
                del_i = i
        start = args[del_i + 1]
        i_row = args[del_i + 2]
        i_col = args[del_i + 3]
        response = delete_row(svc, infile_id, start, i_row, i_col)

def update_doc(doc_id, doc_svc=None, sht_svc=None, dir_svc=None):
    # Get document contents.
    print("Gathering info on existing document...")
    doc_before = get_doc(doc_svc, doc_id)

    # Empty the existing document.
    print("Deleting current contents...")
    end = get_end(doc_before)
    response = delete_all(doc_svc, doc_id, end)

    # Add new content.
    print("Gathering updated content...")
    requests = []
    input_rows = get_sheet(sht_svc, sheet_id)
    photos = get_photos(dir_svc, pics_dir_id)
    abook = create_abook(input_rows, photos)
    rows = create_output_rows(abook)
    table_start_index = 1
    last_cell_index = 16
    rows.reverse()
    for row in rows:
        # Insert new doc table row.
        requests.append(table_insert(table_start_index))
        # Modify table row properties.
        requests.append(table_update_borders(table_start_index + 1))
        # Modify text properties.
        requests.append(table_update_format(table_start_index, last_cell_index))
        # Insert contact data.
        requests.extend(row_data(row, last_cell_index))

    # Execute update.
    print("Gathering updated content...")
    result = send_requests(doc_svc, doc_id, requests)
    print("Done.")

def build_services(services):
    # Initialize variables.
    doc_service = None
    sh_service = None
    dr_service = None
    # Build necessary services.
    print(f"Building service objects for {', '.join(services)}...")
    creds = get_creds(SCOPES)
    if 'docs' in services:
        doc_service = build('docs', 'v1', credentials=creds)
    if 'sheets' in services:
        sh_service = build('sheets', 'v4', credentials=creds)
    if 'drive' in services:
        dr_service = build('drive', 'v3', credentials=creds)
    return {
        'docs': doc_service,
        'sheets': sh_service,
        'drive': dr_service,
    }

def get_creds(scopes):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if Path('token.pickle').is_file():
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

def get_photos(svc, pics_dir_id):
    fields_list = [
        #"id",
        "name",
        "size",
        "webContentLink",
        #"webViewLink",
        #"iconLink",
        #"fullFileExtension",
        #"fileExtension",
    ]
    fields = f"files({', '.join(fields_list)})"
    query = f"'{pics_dir_id}' in parents"
    print("Searching for photos in shared folder...")
    results = svc.files().list(q=query, fields=fields).execute()
    items = results.get('files', [])
    photos_dict = {}
    for i in items:
        name = i["name"].split('.')[0].split('_')[0]
        try:
            photos_dict[name][i["webContentLink"]] = int(i["size"])
        except KeyError:
            photos_dict[name] = {i["webContentLink"]: int(i["size"])}

    return photos_dict

def create_abook(rows, photos):
    # Take data output from spreadsheet and build contacts dictionary.
    abook = {}
    for row in rows[1:]:
        # Pad rows if not enough data.
        full_length = 8
        if len(row) == 1:
            # Heading row (e.g. "Admin/Coordinators"); skip it.
            continue
        elif len(row) < full_length:
            row += [''] * (full_length - len(row))
        # Use data from all columns.
        full_name = f"{row[3]}, {row[4]}"
        abook[full_name] = {}
        for c in range(len(row)):
            abook[full_name][rows[0][c]] = row[c]
        try:
            # List all photo links.
            abook[full_name]['photo'] = photos[full_name]
        except KeyError:
            print(f"Check the spelling for {full_name} in the pictures folder.")
            abook[full_name]['photo'] = None
    return abook

def send_requests(svc, doc_id, requests):
    response = svc.documents().batchUpdate(
        documentId=doc_id,
        body = {'requests': requests}
    ).execute()
    return response

def table_insert(index):
    # Insert a "row" (actually a  table of 2 rows x 3 columns) at the end of the document.
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

def table_update_borders(start):
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

def table_update_format(start_index, end_index):
    request = {
        "updateTextStyle": {
            "range": {
                "startIndex": 12,
                "endIndex": 17,
            },
            "textStyle": {
                "fontSize": {
                    "magnitude": 10,
                    "unit": 'PT'
                }
            },
            "fields": 'fontSize'
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
        # TODO: This is where the rows would be force-sorted.
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

def row_data(row, last_cell_index):
    requests = []
    people_in_row = len(row)

    # Insert text first.
    for i in range(people_in_row):
        # Set position "pos_from_right" so that entries are justified to the
        #   beginning of the row rather than to the end.
        pos_from_right = i + 3 - people_in_row
        # text_index is 12, 14, or 16
        text_index = last_cell_index - 2 * pos_from_right

        # Define contact info variables.
        full_name = f"{row[i]['Last Name']}, {row[i]['First Name']}"
        team = row[i]['Team']
        title = row[i]['Role']
        emails = f"{row[i]['Email']}"
        skype = row[i]['Skype Name']
        phones = f"{row[i]['Phone']}\n"

        # Organize contact data.
        rows = []
        if full_name:
            rows.append(full_name)
        if team and title:
            title_row = f"{team}, {title}"
            rows.append(title_row)
        if emails:
            email_row = f"Email:\n {emails}"
            rows.append(email_row)
        if skype:
            skype_row = f"Skype:\n {skype}"
            rows.append(skype_row)
        if phones:
            phone_row = f"Phone:\n {phones}"
            rows.append(phone_row)
        text = '\n'.join(rows)
        requests.append({
            'insertText': {
                "text": text,
                "location": {
                    "index": text_index
                }
            }
        })

    # Insert photos last.
    for i in range(people_in_row):
        pos_from_right = i + 3 - people_in_row
        # photo_index is 5, 7, or 9
        photo_index = last_cell_index - 7 - 2 * pos_from_right
        photo_links = row[i]['photo']

        # Select largest available photo.
        sizes = [v for v in photo_links.values()]
        sizes.sort(reverse=True)
        photo_link = None
        for link, size in photo_links.items():
            if sizes[0] == size:
                photo_link = link

        if not photo_link:
            print(f"No photo for {row[i]['Last Name']}, {row[i]['First Name']}.")
            # Use placeholder image if photo isn't found.
            photo_link = 'https://drive.google.com/uc?id=1JE7lhkcRWPf0yrasVHu9S0qPWidsktvy&export=download'
        requests.append({
            'insertInlineImage': {
                'location': {
                    'index': photo_index
                },
                'uri': photo_link,
                'objectSize': {
                    'height': {
                        'magnitude': 140,
                        'unit': 'PT'
                    },
                    'width': {
                        'magnitude': 140,
                        'unit': 'PT'
                    }
                }
            }
        })
    return requests

def add_photo(link, index):
    requests = [{
        'insertInlineImage': {
            'location': {
                'index': 1
            },
            'uri':
                link,
            'objectSize': {
                'height': {
                    'magnitude': 50,
                    'unit': 'PT'
                },
                'width': {
                    'magnitude': 50,
                    'unit': 'PT'
                }
            }
        }
    }]
    return requests


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
    requests = [{
        "deleteContentRange": {
            "range": {
                "startIndex": start,
                "endIndex": end,
            }
        }
    }]
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
    """Creates a Google Doc file that merges together photos and contact info
    gleaned from other shared Drive items."""
    do_cmdline(sys.argv, infile_id=template_id, outfile_id=output_id)


if __name__ == '__main__':
    main()
