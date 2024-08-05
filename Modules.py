from Google import Create_Service
from googleapiclient.discovery import build

CLIENT_SECRET_FILE = "Client.json"
API_NAME = "drive"
API_VERSION = "v3"
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents"
]

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
docs_service = build('docs', 'v1', credentials=service._http.credentials)

def file_exists(service, folder_id, file_name):
    """Check if a file with the given name exists in the specified folder."""
    query = f"'{folder_id}' in parents and name='{file_name}' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    if items:
        return items[0]  # Return the file metadata if it exists
    return None

def create_google_doc(service, folder_id, title):
    """Create a new Google Doc in the specified folder."""
    file_metadata = {
        'name': title,
        'mimeType': 'application/vnd.google-apps.document',
        'parents': [folder_id]
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file['id']

def add_text_to_google_doc(docs_service, doc_id, feedback, header):
    """Add text to a Google Doc with the first line as a header."""
    # Get the current content of the document
    document = docs_service.documents().get(documentId=doc_id).execute()

    content_length = 0 # sum([len(element.get('paragraph', {}).get('elements', [{}])[0].get('textRun', {}).get('content', '')) for element in document.get('body').get('content')])
    body = document.get('body')
    content_length = 0
    
    # Calculate the current length of the document's content
    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            for elem in elements:
                if 'textRun' in elem:
                    content_length += len(elem.get('textRun').get('content', '')) 


    requests = [
        {
            'insertText': {
                'location': {
                    'index': content_length,
                },
                'text': header + '\n' + feedback
            }
        },
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': content_length,
                    'endIndex': content_length + len(header)
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {
                        'magnitude': 14,
                        'unit': 'PT',
                    },
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0.7,  
                                'green': 0.0,
                                'blue': 0.5 
                            }
                        }
                    }
                },
                'fields': 'bold,fontSize,foregroundColor'
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()