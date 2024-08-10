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
        return items[0]
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

def content_length_calculater(document):
    content_length = 0
    body = document.get('body', {})

    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            for elem in elements:
                if 'textRun' in elem:
                    text_run = elem.get('textRun')
                    text = text_run.get('content', '')
                    text_length = len(text)
                    
                    # Consider font size (default to 10 if not set)
                    font_size = text_run.get('textStyle', {}).get('fontSize', {}).get('magnitude', 10)
                    
                    # Double the length if the text is bold
                    if text_run.get('textStyle', {}).get('bold', False):
                        text_length *= 2
                    
                    # Add additional length if the text is italic (e.g., increase by 50%)
                    if text_run.get('textStyle', {}).get('italic', False):
                        text_length = int(text_length * 1.5)
                    
                    # Adjust length based on font size (relative to size 10)
                    text_length = int(text_length * (font_size / 10))
                    
                    content_length += text_length
        
        # Consider embedded objects like images
        if 'inlineObjectElement' in element:
            # Arbitrary length assigned for images or other inline objects
            content_length += 100  # Adjust as needed
            
        # Consider table cells (could be expanded to include more details)
        if 'table' in element:
            table = element.get('table')
            rows = table.get('tableRows', [])
            for row in rows:
                cells = row.get('tableCells', [])
                for cell in cells:
                    for content in cell.get('content', []):
                        if 'paragraph' in content:
                            elements = content.get('paragraph').get('elements', [])
                            for elem in elements:
                                if 'textRun' in elem:
                                    text_run = elem.get('textRun')
                                    text = text_run.get('content', '')
                                    text_length = len(text)
                                    
                                    # Consider font size, bold, and italic formatting in table cells
                                    font_size = text_run.get('textStyle', {}).get('fontSize', {}).get('magnitude', 10)
                                    if text_run.get('textStyle', {}).get('bold', False):
                                        text_length *= 2
                                    if text_run.get('textStyle', {}).get('italic', False):
                                        text_length = int(text_length * 1.5)
                                    text_length = int(text_length * (font_size / 10))
                                    
                                    content_length += text_length
    return content_length


def add_text_to_google_doc_textFormat(docs_service, doc_id, feedback, header):
    document = docs_service.documents().get(documentId=doc_id).execute()

    content_length = 0 
    body = document.get('body')
    content_length = 0
    
    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            #print(elements)
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
                        'magnitude': 11,
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
                    },
                    'backgroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0.9,  
                                'green': 0.9,
                                'blue': 0.9
                            }
                        }
                    },
                    'weightedFontFamily': {
                        'fontFamily': 'Times New Roman',
                        'weight': 400 
                    }
              
                },
                'fields': 'bold,fontSize,foregroundColor,backgroundColor,weightedFontFamily'
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

def add_text_to_google_doc_cellFormat(docs_service, doc_id, feedback, header):
    document = docs_service.documents().get(documentId=doc_id).execute()

    content_length = 0
    body = document.get('body')
    content_length = 0
    
    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            #print(elements)
            for elem in elements:
                if 'textRun' in elem:
                    content_length += len(elem.get('textRun').get('content', ''))
            
    requests = [
        {
            'insertText': {
                'location': {
                    'index': content_length  
                },
                'text': header + '\n'
            }
        },
        {
            'insertTable': {
                'rows': 1,
                'columns': 1,
                'location': {
                    'index': content_length + len(header) 
                }
            }
        },
        {
            'insertText': {
                'location': {
                    'index': content_length + len(header) + 4 
                },
                'text': feedback
            }
        }, 
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': content_length + len(header) + 4,
                    'endIndex': content_length + len(header) + len(feedback) + 4
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {
                        'magnitude': 11,
                        'unit': 'PT',
                    },
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0,  
                                'green': 0,
                                'blue': 0 
                            }
                        }
                    },
                    'weightedFontFamily': {
                        'fontFamily': 'Times New Roman',
                        'weight': 400 
                    }
                },
                'fields': 'bold,fontSize,foregroundColor,weightedFontFamily'
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()


def add_text_to_google_doc_afterDeterminedHeader_textFormat(docs_service, doc_id, feedback, score, determinedHeader):
    document = docs_service.documents().get(documentId=doc_id).execute()
    
    content_length = 0 
    body = document.get('body')
    content_length = 0
    
    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            #print(elements)
            for elem in elements:
                if 'textRun' in elem:
                    content_length += len(elem.get('textRun').get('content', '')) 
                    if determinedHeader in elem.get('textRun').get('content', ''):
                        break
            if determinedHeader in elem.get('textRun').get('content', ''):
               break
    
    requests = [
        {
            'insertText': {
                'location': {
                    'index': content_length + 1
                },
                'text': score + '\n'
            }
        },
        {
            'insertText': {
                'location': {
                    'index': content_length + len(score) + 2,
                },
                'text':  feedback
            }
        },
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': content_length + len(score) + 2,
                    'endIndex': content_length + len(score) + len(feedback) + 2
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {
                        'magnitude': 11,
                        'unit': 'PT',
                    },
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0,  
                                'green': 0,
                                'blue': 0 
                            }
                        }
                    },
                    'weightedFontFamily': {
                        'fontFamily': 'Times New Roman',
                        'weight': 400 
                    }
                },
                'fields': 'bold,fontSize,foregroundColor,weightedFontFamily'
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

  
def add_text_to_google_doc_afterDeterminedHeader_cellFormat(docs_service, doc_id, feedback, score, determinedHeader):
    document = docs_service.documents().get(documentId=doc_id).execute()

    content_length = 0 
    body = document.get('body')
    content_length = 0
    
    for element in body.get('content', []):
        if 'paragraph' in element:
            elements = element.get('paragraph').get('elements', [])
            #print(elements)
            for elem in elements:
                if 'textRun' in elem:
                    content_length += len(elem.get('textRun').get('content', '')) 
                    if determinedHeader in elem.get('textRun').get('content', ''):
                        break
            if determinedHeader in elem.get('textRun').get('content', ''):
                break
    
    requests = [
        {
            'insertText': {
                'location': {
                    'index': content_length + 1
                },
                'text': score + "\n"
            }
        },
        {
            'insertTable': {
                'rows': 1,
                'columns': 1,
                'location': {
                    'index': content_length + len(score) + 1
                }
            }
        },
        {
            'insertText': {
                'location': {
                    'index': content_length + len(score) + 1 + 4 
                },
                'text': feedback
            }
        }, 
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': content_length + len(score) + 1 + 4,
                    'endIndex': content_length + len(score) + len(feedback) + 1 + 4
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {
                        'magnitude': 11,
                        'unit': 'PT',
                    },
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'red': 0,  
                                'green': 0,
                                'blue': 0 
                            }
                        }
                    },
                    'weightedFontFamily': {
                        'fontFamily': 'Times New Roman',
                        'weight': 400 
                    }
                },
                'fields': 'bold,fontSize,foregroundColor,weightedFontFamily'
            }
        }
    ]

    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()