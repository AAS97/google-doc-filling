from __future__ import print_function

import json
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

# The ID of a sample document.
TEMPLATE_ID = 'placeholder' ## Replace with appropriate id from document url


def duplicateDocument(templateId, title = "New Document"):	
    # Retrieve the documents contents from the Docs service.
    doc_resp = drive_service.files().copy(fileId=templateId, body = {"name": title}).execute()
    docId = doc_resp.get("id");
    return docId

def replace_text(document_id, replacements):
    """Replaces the text in existing named ranges."""

    # Fetch the document to determine the current indexes of the named ranges.
    document = doc_service.documents().get(documentId=document_id).execute()

    # Create a sequence of requests for each range.
    requests = []
    for replacement in replacements:
        # Delete all the content in the existing range.
        requests.append({"replaceAllText": {
        "containsText": { 
          "matchCase": True,
          "text": replacement["tag"], # The text to search for in the document.
        },
        "replaceText": replacement["text"], # The text that will replace the matched text.
      }})

    # Make a batchUpdate request to apply the changes, ensuring the document
    # hasn't changed since we fetched it.
    body = {
        'requests': requests,
        'writeControl': {
            'requiredRevisionId': document.get('revisionId')
        }
    }
    doc_service.documents().batchUpdate(documentId=document_id, body=body).execute()

def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    global creds
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())


    try:
        global doc_service, drive_service
        doc_service = build('docs', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials = creds)
        
        documentId = duplicateDocument(templateId=TEMPLATE_ID, title="New Document")
        with open("replacements.json", "r") as file : 
            replacements = json.load(file)
        # document = doc_service.documents().get(documentId=documentId).execute()
        replace_text(documentId, replacements)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
