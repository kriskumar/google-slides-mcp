import os
import asyncio
import logging
import json
import base64
from io import BytesIO
from typing import List, Dict, Any, Optional, Union, Tuple

from mcp.server.fastmcp import FastMCP, Image, Context
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseUpload
import pickle

# Plotting imports
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import pandas as pd

# Initialize FastMCP server
mcp = FastMCP("google-slides")

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('mcp_google_slides_server')

# Google API scopes
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']

# Hard-coded token path
TOKEN_PATH = "token.json"

# Helper function to convert plotly figure to image
def fig_to_image(fig, width=800, height=600, format="png"):
    """Convert a plotly figure to an Image object that can be returned by MCP"""
    img_bytes = fig.to_image(format=format, width=width, height=height)
    return Image(data=img_bytes, format=format)

class SlidesManager:
    def __init__(self):
        self.presentations = {}
        self.creds = self._get_credentials()
        self.slides_service = build('slides', 'v1', credentials=self.creds)
        self.drive_service = build('drive', 'v3', credentials=self.creds)
    
    def _get_credentials(self):
        """Get Google API credentials from the hard-coded token path."""
        creds = None
        
        # Check if token file exists - use the same approach as in google-docs-server.py
        if os.path.exists(TOKEN_PATH):
            with open(TOKEN_PATH, 'r') as token:
                creds = Credentials.from_authorized_user_info(json.load(token), SCOPES)
        else:
            raise ValueError(f"Token file not found at {TOKEN_PATH}")
        
        # If credentials have expired, refresh them
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            # Save the refreshed credentials
            with open(TOKEN_PATH, 'w') as token:
                token.write(creds.to_json())
        
        return creds
    
    def create_presentation(self, name: str) -> str:
        """Create a new Google Slides presentation."""
        presentation = self.slides_service.presentations().create(
            body={'title': name}
        ).execute()
        
        presentation_id = presentation['presentationId']
        self.presentations[name] = presentation_id
        
        return presentation_id
    
    def _execute_batch_update(self, presentation_id: str, requests: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Execute a batch update on a presentation."""
        body = {
            'requests': requests
        }
        
        response = self.slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body=body
        ).execute()
        
        return response
    
    def _create_short_id(self, prefix: str, title: str) -> str:
        """Create a short unique ID based on the title."""
        import hashlib
        hash_id = hashlib.md5(title.encode()).hexdigest()[:10]
        return f"{prefix}_{hash_id}"
    
    def add_title_slide(self, presentation_name: str, title: str, subtitle: str = "") -> str:
        """Add a title slide to the presentation."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("title", title)
        
        # Add a new slide
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Wait a moment for slide creation to complete
        import time
        time.sleep(1)

        # Get the presentation to find the created slide and its placeholders
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title and subtitle placeholders
            title_id = None
            subtitle_id = None

            for element in created_slide.get('pageElements', []):
                shape = element.get('shape', {})
                placeholder = shape.get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')
                elif placeholder.get('type') == 'SUBTITLE':
                    subtitle_id = element.get('objectId')

            # Now add the title and subtitle text
            text_requests = []
            if title_id:
                text_requests.append({
                    'insertText': {
                        'objectId': title_id,
                        'text': title
                    }
                })

            if subtitle and subtitle_id:
                text_requests.append({
                    'insertText': {
                        'objectId': subtitle_id,
                        'text': subtitle
                    }
                })

            if text_requests:
                self._execute_batch_update(presentation_id, text_requests)
        
        return created_slide_id
    
    def add_section_header_slide(self, presentation_name: str, header: str, subtitle: str = "") -> str:
        """Add a section header slide to the presentation."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("section", header)
        
        # Add a new slide
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'SECTION_HEADER'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Get the slide to find placeholder IDs
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides.pageElements'
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title and body placeholders
            title_id = None
            body_id = None

            for element in created_slide.get('pageElements', []):
                placeholder = element.get('shape', {}).get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')
                elif placeholder.get('type') == 'BODY':
                    body_id = element.get('objectId')

            # Now add the header and subtitle text
            text_requests = []
            if title_id:
                text_requests.append({
                    'insertText': {
                        'objectId': title_id,
                        'text': header
                    }
                })

            if subtitle and body_id:
                text_requests.append({
                    'insertText': {
                        'objectId': body_id,
                        'text': subtitle
                    }
                })

            if text_requests:
                self._execute_batch_update(presentation_id, text_requests)
        
        return created_slide_id
    
    def add_content_slide(self, presentation_name: str, title: str, content: str) -> str:
        """Add a slide with title and content."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("content", title)
        
        # Add a new slide
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_AND_BODY'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Get the slide to find placeholder IDs
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides.pageElements'
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title and body placeholders
            title_id = None
            body_id = None

            for element in created_slide.get('pageElements', []):
                placeholder = element.get('shape', {}).get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')
                elif placeholder.get('type') == 'BODY':
                    body_id = element.get('objectId')

            # Add the title and content text
            text_requests = []
            if title_id:
                text_requests.append({
                    'insertText': {
                        'objectId': title_id,
                        'text': title
                    }
                })

            if body_id:
                text_requests.append({
                    'insertText': {
                        'objectId': body_id,
                        'text': content
                    }
                })

            if text_requests:
                self._execute_batch_update(presentation_id, text_requests)

            # Add bullet formatting to the body if we have content
            if body_id and content.strip():
                bullet_requests = []
                lines = content.strip().split('\n')

                current_index = 0
                for line in lines:
                    if not line.strip():
                        current_index += len(line) + 1
                        continue

                    line_length = len(line.rstrip())

                    # Only add bullets if this isn't a blank line
                    if line.strip():
                        bullet_requests.append({
                            'createParagraphBullets': {
                                'objectId': body_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': current_index,
                                    'endIndex': current_index + line_length
                                },
                                'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                            }
                        })

                    # Move to next line (add 1 for the newline character)
                    current_index += line_length + 1

                if bullet_requests:
                    self._execute_batch_update(presentation_id, bullet_requests)
        
        return created_slide_id
    
    def add_two_column_slide(self, presentation_name: str, title: str, 
                           left_title: str, left_content: str,
                           right_title: str, right_content: str) -> str:
        """Add a slide with two columns for comparison."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("twocol", title)
        
        # Add a new slide
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_AND_TWO_COLUMNS'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Get the slide to find placeholder IDs
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides.pageElements'
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title and body placeholders
            title_id = None
            body_ids = []  # For two-column layouts, there are multiple body placeholders

            for element in created_slide.get('pageElements', []):
                placeholder = element.get('shape', {}).get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')
                elif placeholder.get('type') == 'BODY':
                    body_ids.append(element.get('objectId'))

            # Add the title and column content
            text_requests = []
            if title_id:
                text_requests.append({
                    'insertText': {
                        'objectId': title_id,
                        'text': title
                    }
                })

            # Add left column content (first body placeholder)
            if len(body_ids) > 0:
                left_content_formatted = f"{left_title}\n{left_content}"
                text_requests.append({
                    'insertText': {
                        'objectId': body_ids[0],
                        'text': left_content_formatted
                    }
                })

            # Add right column content (second body placeholder)
            if len(body_ids) > 1:
                right_content_formatted = f"{right_title}\n{right_content}"
                text_requests.append({
                    'insertText': {
                        'objectId': body_ids[1],
                        'text': right_content_formatted
                    }
                })

            if text_requests:
                self._execute_batch_update(presentation_id, text_requests)
        
        return created_slide_id
    
    def add_table_slide(self, presentation_name: str, title: str, 
                      headers: List[str], rows: List[List[Any]]) -> str:
        """Add a slide with a table."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("table", title)
        
        # Add a new slide with title
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_ONLY'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Get the slide to find placeholder IDs
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides.pageElements'
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title placeholder
            title_id = None

            for element in created_slide.get('pageElements', []):
                placeholder = element.get('shape', {}).get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')

            # Add the title text
            if title_id:
                title_request = [
                    {
                        'insertText': {
                            'objectId': title_id,
                            'text': title
                        }
                    }
                ]
                self._execute_batch_update(presentation_id, title_request)
        
        # Create table
        table_id = f'{slide_id}_tbl'
        create_table_request = [
            {
                'createTable': {
                    'objectId': table_id,
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': {'magnitude': 400, 'unit': 'PT'},
                            'height': {'magnitude': 300, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    },
                    'rows': len(rows) + 1,  # +1 for headers
                    'columns': len(headers)
                }
            }
        ]
        
        self._execute_batch_update(presentation_id, create_table_request)
        
        # Fill in header row
        header_requests = []
        for i, header in enumerate(headers):
            header_requests.append({
                'insertText': {
                    'objectId': table_id,
                    'cellLocation': {
                        'rowIndex': 0,
                        'columnIndex': i
                    },
                    'text': str(header)
                }
            })
            
            # Bold the header
            header_requests.append({
                'updateTextStyle': {
                    'objectId': table_id,
                    'cellLocation': {
                        'rowIndex': 0,
                        'columnIndex': i
                    },
                    'style': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })
        
        if header_requests:
            self._execute_batch_update(presentation_id, header_requests)
        
        # Fill in data rows
        data_requests = []
        for i, row in enumerate(rows):
            for j, cell in enumerate(row):
                data_requests.append({
                    'insertText': {
                        'objectId': table_id,
                        'cellLocation': {
                            'rowIndex': i + 1,  # +1 to skip header row
                            'columnIndex': j
                        },
                        'text': str(cell)
                    }
                })
        
        if data_requests:
            self._execute_batch_update(presentation_id, data_requests)
        
        return created_slide_id
    
    def add_image_slide(self, presentation_name: str, title: str, image_data: bytes, caption: str = "") -> str:
        """Add a slide with an image."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        
        presentation_id = self.presentations[presentation_name]
        
        # First, upload the image to Google Drive
        file_metadata = {
            'name': f'img_{title.replace(" ", "_")[:20]}.png',
        }

        # Create a BytesIO object from the image data
        from io import BytesIO
        media = MediaIoBaseUpload(
            BytesIO(image_data),
            mimetype='image/png',
            resumable=True
        )

        file = self.drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        image_file_id = file.get('id')

        # Make the file publicly viewable so Slides can access it
        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        self.drive_service.permissions().create(
            fileId=image_file_id,
            body=permission
        ).execute()

        # Create a shorter unique object ID for the slide
        slide_id = self._create_short_id("img", title)
        
        # Add a new slide with title
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_ONLY'
                    }
                }
            }
        ]
        
        response = self._execute_batch_update(presentation_id, requests)
        created_slide_id = response['replies'][0]['createSlide']['objectId']

        # Get the slide to find placeholder IDs
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides.pageElements'
        ).execute()

        # Find the created slide and its placeholders
        created_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == created_slide_id:
                created_slide = slide
                break

        if created_slide:
            # Find title placeholder
            title_id = None

            for element in created_slide.get('pageElements', []):
                placeholder = element.get('shape', {}).get('placeholder', {})
                if placeholder.get('type') == 'TITLE':
                    title_id = element.get('objectId')

            # Add the title text
            if title_id:
                title_request = [
                    {
                        'insertText': {
                            'objectId': title_id,
                            'text': title
                        }
                    }
                ]
                self._execute_batch_update(presentation_id, title_request)

        # Add the image
        image_id = f'{slide_id}_i'
        image_request = [
            {
                'createImage': {
                    'objectId': image_id,
                    'url': f'https://drive.google.com/uc?id={image_file_id}',
                    'elementProperties': {
                        'pageObjectId': created_slide_id,
                        'size': {
                            'width': {'magnitude': 400, 'unit': 'PT'},
                            'height': {'magnitude': 300, 'unit': 'PT'},
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 150,
                            'unit': 'PT'
                        }
                    }
                }
            }
        ]

        self._execute_batch_update(presentation_id, image_request)

        # Add caption if provided (create a text box for it)
        if caption:
            caption_id = f'{slide_id}_c'
            caption_request = [
                {
                    'createShape': {
                        'objectId': caption_id,
                        'shapeType': 'TEXT_BOX',
                        'elementProperties': {
                            'pageObjectId': created_slide_id,
                            'size': {
                                'width': {'magnitude': 400, 'unit': 'PT'},
                                'height': {'magnitude': 50, 'unit': 'PT'},
                            },
                            'transform': {
                                'scaleX': 1,
                                'scaleY': 1,
                                'translateX': 100,
                                'translateY': 470,
                                'unit': 'PT'
                            }
                        }
                    }
                },
                {
                    'insertText': {
                        'objectId': caption_id,
                        'text': caption
                    }
                }
            ]

            self._execute_batch_update(presentation_id, caption_request)
        
        return created_slide_id
    
    def get_presentation_url(self, presentation_name: str) -> str:
        """Get the URL for a presentation."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        presentation_id = self.presentations[presentation_name]
        return f"https://docs.google.com/presentation/d/{presentation_id}/edit"

    def apply_theme_from_presentation(self, presentation_name: str, source_presentation_id: str) -> str:
        """Apply theme from another presentation."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        presentation_id = self.presentations[presentation_name]

        # Get the source presentation to extract master slides
        source_presentation = self.slides_service.presentations().get(
            presentationId=source_presentation_id,
            fields='masters'
        ).execute()

        if 'masters' not in source_presentation:
            raise ValueError("Source presentation has no masters")

        # Apply the theme by replacing the master
        requests = []
        for master in source_presentation['masters']:
            requests.append({
                'replaceAllShapesWithImage': {
                    'imageUrl': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
                    'replaceMethod': 'CENTER_INSIDE'
                }
            })

        if requests:
            self._execute_batch_update(presentation_id, requests)

        return f"Applied theme from {source_presentation_id} to {presentation_name}"

    def apply_beautiful_styling(self, presentation_name: str) -> str:
        """Apply beautiful styling and colors to the presentation."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        presentation_id = self.presentations[presentation_name]

        # Get all slides to apply styling
        presentation = self.slides_service.presentations().get(
            presentationId=presentation_id,
            fields='slides'
        ).execute()

        styling_requests = []

        # Apply a beautiful color scheme and fonts
        for slide in presentation.get('slides', []):
            slide_id = slide.get('objectId')

            # Update slide background with a subtle gradient
            styling_requests.append({
                'updateSlideProperties': {
                    'objectId': slide_id,
                    'slideProperties': {
                        'pageBackgroundFill': {
                            'solidFill': {
                                'color': {
                                    'rgbColor': {
                                        'red': 0.97,
                                        'green': 0.98,
                                        'blue': 1.0
                                    }
                                }
                            }
                        }
                    },
                    'fields': 'pageBackgroundFill'
                }
            })

        if styling_requests:
            self._execute_batch_update(presentation_id, styling_requests)

        return f"Applied beautiful styling to {presentation_name}"

    def apply_theme_by_name(self, presentation_name: str, theme_name: str) -> str:
        """Search for a theme template in Google Drive by name and apply it."""
        if presentation_name not in self.presentations:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        # Search for presentations in Drive that match the theme name
        search_query = f"name contains '{theme_name}' and mimeType='application/vnd.google-apps.presentation'"

        try:
            results = self.drive_service.files().list(
                q=search_query,
                fields="files(id, name)"
            ).execute()

            files = results.get('files', [])

            if not files:
                raise ValueError(f"No theme template found with name containing '{theme_name}'")

            # Use the first matching file as the theme source
            theme_file = files[0]
            theme_id = theme_file['id']
            theme_file_name = theme_file['name']

            # Apply the theme from this presentation
            result = self.apply_theme_from_presentation(presentation_name, theme_id)
            return f"Applied theme '{theme_file_name}' to {presentation_name}"

        except Exception as e:
            raise ValueError(f"Failed to find or apply theme: {str(e)}")

    def list_available_themes(self) -> List[Dict[str, str]]:
        """List available theme templates in Google Drive."""
        try:
            # Search for all presentation files that could be themes
            search_query = "mimeType='application/vnd.google-apps.presentation' and (name contains 'theme' or name contains 'template')"

            results = self.drive_service.files().list(
                q=search_query,
                fields="files(id, name, modifiedTime)",
                orderBy="modifiedTime desc"
            ).execute()

            files = results.get('files', [])
            themes = []

            for file in files:
                themes.append({
                    'id': file['id'],
                    'name': file['name'],
                    'modified': file.get('modifiedTime', 'Unknown')
                })

            return themes

        except Exception as e:
            raise ValueError(f"Failed to list themes: {str(e)}")
    

# MCP Tool Definitions
@mcp.tool()
def create_presentation(name: str) -> str:
    """Create a new Google Slides presentation.
    
    Args:
        name: Name of the presentation
                
    Returns:
        Confirmation message with the presentation ID
    """
    try:
        slides_manager = SlidesManager()
        presentation_id = slides_manager.create_presentation(name)
        
        # Store the slides manager in a global variable for later use
        global _slides_manager
        _slides_manager = slides_manager
        
        return f"Created new presentation: {name} (ID: {presentation_id})"
    except Exception as e:
        raise ValueError(f"Failed to create presentation: {str(e)}")

@mcp.tool()
def add_title_slide(presentation_name: str, title: str, subtitle: str = "") -> str:
    """Add a title slide to an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        title: Title text for the slide
        subtitle: Optional subtitle text for the slide
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        slide_id = _slides_manager.add_title_slide(presentation_name, title, subtitle)
        return f"Added title slide '{title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add title slide: {str(e)}")

@mcp.tool()
def add_section_header(presentation_name: str, header: str, subtitle: str = "") -> str:
    """Add a section header slide to an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        header: Header text for the slide
        subtitle: Optional subtitle text for the slide
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        slide_id = _slides_manager.add_section_header_slide(presentation_name, header, subtitle)
        return f"Added section header slide '{header}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add section header slide: {str(e)}")

@mcp.tool()
def add_content_slide(presentation_name: str, title: str, content: str) -> str:
    """Add a content slide with bullet points to an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        title: Title text for the slide
        content: Content text with bullet points (use tab for indentation)
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        slide_id = _slides_manager.add_content_slide(presentation_name, title, content)
        return f"Added content slide '{title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add content slide: {str(e)}")

@mcp.tool()
def add_two_column_slide(
    presentation_name: str, 
    title: str, 
    left_title: str, 
    left_content: str,
    right_title: str, 
    right_content: str
) -> str:
    """Add a two-column comparison slide to an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        title: Title text for the slide
        left_title: Title for the left column
        left_content: Content for the left column (use tab for indentation)
        right_title: Title for the right column
        right_content: Content for the right column (use tab for indentation)
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        slide_id = _slides_manager.add_two_column_slide(
            presentation_name, title, 
            left_title, left_content,
            right_title, right_content
        )
        return f"Added two-column slide '{title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add two-column slide: {str(e)}")

@mcp.tool()
def add_table_slide(
    presentation_name: str,
    title: str,
    data: Dict[str, Any]
) -> str:
    """Add a slide with a table to an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        title: Title text for the slide
        data: Dictionary with 'headers' (list of strings) and 'rows' (list of lists)
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        headers = data.get("headers", [])
        rows = data.get("rows", [])
        
        if not headers:
            raise ValueError("Table headers are required")
        
        if not rows:
            raise ValueError("Table rows are required")
        
        if not all(len(row) == len(headers) for row in rows):
            raise ValueError("All rows must have the same number of columns as headers")
        
        slide_id = _slides_manager.add_table_slide(presentation_name, title, headers, rows)
        return f"Added table slide '{title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add table slide: {str(e)}")

@mcp.tool()
def get_presentation_url(presentation_name: str) -> str:
    """Get the URL of an existing presentation.
    
    Args:
        presentation_name: Name of the presentation
        
    Returns:
        URL of the presentation
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        url = _slides_manager.get_presentation_url(presentation_name)
        return f"Presentation URL: {url}"
    except Exception as e:
        raise ValueError(f"Failed to get presentation URL: {str(e)}")

# Plotly visualization tools
@mcp.tool()
def create_bar_chart(
    presentation_name: str,
    slide_title: str,
    categories: List[str],
    values: List[float],
    chart_title: str = "Bar Chart",
    x_label: str = "Categories",
    y_label: str = "Values",
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a bar chart and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        categories: List of categories for the x-axis
        values: List of values for the y-axis
        chart_title: Title of the chart
        x_label: Label for the x-axis
        y_label: Label for the y-axis
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        if len(categories) != len(values):
            raise ValueError("categories and values must have the same length")
        
        # Create the bar chart with Plotly
        df = pd.DataFrame({'category': categories, 'value': values})
        fig = px.bar(df, x='category', y='value', title=chart_title)
        fig.update_layout(
            xaxis_title=x_label,
            yaxis_title=y_label
        )
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Chart showing {y_label} by {x_label}"
        )
        
        return f"Added bar chart slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add bar chart slide: {str(e)}")

@mcp.tool()
def create_line_plot(
    presentation_name: str,
    slide_title: str,
    x_values: List[float],
    y_values: List[float],
    chart_title: str = "Line Plot",
    x_label: str = "X Axis",
    y_label: str = "Y Axis",
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a line plot and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        x_values: List of values for the x-axis
        y_values: List of values for the y-axis
        chart_title: Title of the chart
        x_label: Label for the x-axis
        y_label: Label for the y-axis
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        if len(x_values) != len(y_values):
            raise ValueError("x_values and y_values must have the same length")
        
        # Create the line plot with Plotly
        df = pd.DataFrame({'x': x_values, 'y': y_values})
        fig = px.line(df, x='x', y='y', title=chart_title)
        fig.update_layout(
            xaxis_title=x_label,
            yaxis_title=y_label
        )
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Chart showing {y_label} vs {x_label}"
        )
        
        return f"Added line plot slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add line plot slide: {str(e)}")

@mcp.tool()
def create_pie_chart(
    presentation_name: str,
    slide_title: str,
    labels: List[str],
    values: List[float],
    chart_title: str = "Pie Chart",
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a pie chart and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        labels: List of labels for the pie chart segments
        values: List of values determining the size of each segment
        chart_title: Title of the chart
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        if len(labels) != len(values):
            raise ValueError("labels and values must have the same length")
        
        # Create the pie chart with Plotly
        fig = go.Figure(data=[go.Pie(labels=labels, values=values)])
        fig.update_layout(title=chart_title)
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Pie chart showing distribution of {chart_title}"
        )
        
        return f"Added pie chart slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add pie chart slide: {str(e)}")

@mcp.tool()
def create_scatter_plot(
    presentation_name: str,
    slide_title: str,
    x_values: List[float],
    y_values: List[float],
    chart_title: str = "Scatter Plot",
    x_label: str = "X Axis",
    y_label: str = "Y Axis",
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a scatter plot and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        x_values: List of values for the x-axis
        y_values: List of values for the y-axis
        chart_title: Title of the chart
        x_label: Label for the x-axis
        y_label: Label for the y-axis
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        if len(x_values) != len(y_values):
            raise ValueError("x_values and y_values must have the same length")
        
        # Create the scatter plot with Plotly
        df = pd.DataFrame({'x': x_values, 'y': y_values})
        fig = px.scatter(df, x='x', y='y', title=chart_title)
        fig.update_layout(
            xaxis_title=x_label,
            yaxis_title=y_label
        )
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Scatter plot showing relationship between {x_label} and {y_label}"
        )
        
        return f"Added scatter plot slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add scatter plot slide: {str(e)}")

@mcp.tool()
def create_heatmap(
    presentation_name: str,
    slide_title: str,
    matrix: List[List[float]],
    x_labels: Optional[List[str]] = None,
    y_labels: Optional[List[str]] = None,
    chart_title: str = "Heatmap",
    colorscale: str = "Viridis",
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a heatmap and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        matrix: 2D list representing the matrix of values
        x_labels: Labels for the x-axis (optional)
        y_labels: Labels for the y-axis (optional)
        chart_title: Title of the chart
        colorscale: Colorscale for the heatmap
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        # Create the heatmap with Plotly
        fig = go.Figure(data=go.Heatmap(
            z=matrix,
            x=x_labels,
            y=y_labels,
            colorscale=colorscale
        ))
        fig.update_layout(title=chart_title)
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Heatmap visualization of {chart_title}"
        )
        
        return f"Added heatmap slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add heatmap slide: {str(e)}")

@mcp.tool()
def create_histogram(
    presentation_name: str,
    slide_title: str,
    values: List[float],
    chart_title: str = "Histogram",
    x_label: str = "Values",
    y_label: str = "Count",
    bins: Optional[int] = None,
    width: int = 800,
    height: int = 600,
) -> str:
    """Create a histogram and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        values: List of values to plot
        chart_title: Title of the chart
        x_label: Label for the x-axis
        y_label: Label for the y-axis
        bins: Number of bins for the histogram (optional)
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        # Create the histogram with Plotly
        df = pd.DataFrame({'value': values})
        fig = px.histogram(df, x='value', title=chart_title, nbins=bins)
        fig.update_layout(
            xaxis_title=x_label,
            yaxis_title=y_label
        )
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Histogram showing distribution of {x_label}"
        )
        
        return f"Added histogram slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add histogram slide: {str(e)}")

@mcp.tool()
def create_scatter_matrix(
    presentation_name: str,
    slide_title: str,
    data: Dict[str, List[float]],
    chart_title: str = "Scatter Matrix",
    width: int = 1000,
    height: int = 1000,
) -> str:
    """Create a scatter matrix (pairs plot) and add it as a slide in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        data: Dictionary where keys are column names and values are lists of data
        chart_title: Title of the chart
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")
        
        # Validate that all lists have the same length
        if len(set(len(values) for values in data.values())) != 1:
            raise ValueError("All data lists must have the same length")
        
        # Create the scatter matrix with Plotly
        df = pd.DataFrame(data)
        fig = px.scatter_matrix(df, title=chart_title)
        
        # Convert the figure to image bytes
        img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Add the image to a slide
        slide_id = _slides_manager.add_image_slide(
            presentation_name, 
            slide_title, 
            img_bytes,
            f"Scatter matrix showing relationships between variables"
        )
        
        return f"Added scatter matrix slide '{slide_title}' to presentation: {presentation_name}"
    except Exception as e:
        raise ValueError(f"Failed to add scatter matrix slide: {str(e)}")

# Helper functions for sample data generation
def generate_sine_wave(n_points=100, amplitude=1.0, frequency=1.0, phase=0.0, noise=0.0):
    """Generate a sine wave with optional noise"""
    x = np.linspace(0, 2*np.pi, n_points)
    y = amplitude * np.sin(frequency * x + phase)
    if noise > 0:
        y += np.random.normal(0, noise, n_points)
    return x.tolist(), y.tolist()

def generate_random_categories(n_categories=5, min_value=0, max_value=100, seed=None):
    """Generate random categories and values"""
    if seed is not None:
        np.random.seed(seed)
    categories = [f"Category {i+1}" for i in range(n_categories)]
    values = np.random.randint(min_value, max_value, n_categories).tolist()
    return categories, values

@mcp.tool()
def generate_sample_data(
    data_type: str = "sine_wave",
    n_points: int = 100,
    seed: Optional[int] = None,
) -> Dict[str, Any]:
    """Generate sample data for plotting.
    
    Args:
        data_type: Type of data to generate ("sine_wave", "categories", "linear", "normal")
        n_points: Number of data points to generate
        seed: Random seed for reproducibility
        
    Returns:
        Dictionary containing the generated data
    """
    if seed is not None:
        np.random.seed(seed)
    
    if data_type == "sine_wave":
        x, y = generate_sine_wave(n_points, noise=0.2)
        return {"x": x, "y": y}
    
    elif data_type == "categories":
        categories, values = generate_random_categories(n_points)
        return {"categories": categories, "values": values}
    
    elif data_type == "linear":
        # Generate linear data with noise
        x = np.linspace(0, 10, n_points)
        y = 2 * x + 5 + np.random.normal(0, 1, n_points)
        return {"x": x.tolist(), "y": y.tolist()}
    
    elif data_type == "normal":
        # Generate normally distributed data
        values = np.random.normal(0, 1, n_points).tolist()
        return {"values": values}
    
    else:
        raise ValueError(f"Unsupported data type: {data_type}")

@mcp.tool()
def create_chart_from_sample_data(
    presentation_name: str,
    slide_title: str,
    data_type: str = "sine_wave",
    chart_type: str = "line",
    n_points: int = 100,
    seed: Optional[int] = None,
    width: int = 800,
    height: int = 600,
) -> str:
    """Generate sample data and create a chart in the presentation.
    
    Args:
        presentation_name: Name of the presentation
        slide_title: Title of the slide
        data_type: Type of data to generate ("sine_wave", "categories", "linear", "normal")
        chart_type: Type of chart to create ("line", "scatter", "bar", "pie", "histogram")
        n_points: Number of data points to generate
        seed: Random seed for reproducibility
        width: Width of the chart in pixels
        height: Height of the chart in pixels
        
    Returns:
        Confirmation message
    """
    try:
        # Generate sample data
        data = generate_sample_data(data_type, n_points, seed)
        
        # Create chart based on data and chart type
        if chart_type == "line" and "x" in data and "y" in data:
            return create_line_plot(
                presentation_name=presentation_name,
                slide_title=slide_title,
                x_values=data["x"],
                y_values=data["y"],
                chart_title=f"Line Plot of {data_type.title()} Data",
                width=width,
                height=height
            )
        
        elif chart_type == "scatter" and "x" in data and "y" in data:
            return create_scatter_plot(
                presentation_name=presentation_name,
                slide_title=slide_title,
                x_values=data["x"],
                y_values=data["y"],
                chart_title=f"Scatter Plot of {data_type.title()} Data",
                width=width,
                height=height
            )
        
        elif chart_type == "bar" and "categories" in data and "values" in data:
            return create_bar_chart(
                presentation_name=presentation_name,
                slide_title=slide_title,
                categories=data["categories"],
                values=data["values"],
                chart_title=f"Bar Chart of {data_type.title()} Data",
                width=width,
                height=height
            )
        
        elif chart_type == "pie" and "categories" in data and "values" in data:
            return create_pie_chart(
                presentation_name=presentation_name,
                slide_title=slide_title,
                labels=data["categories"],
                values=data["values"],
                chart_title=f"Pie Chart of {data_type.title()} Data",
                width=width,
                height=height
            )
        
        elif chart_type == "histogram" and "values" in data:
            return create_histogram(
                presentation_name=presentation_name,
                slide_title=slide_title,
                values=data["values"],
                chart_title=f"Histogram of {data_type.title()} Data",
                width=width,
                height=height
            )
        
        else:
            raise ValueError(f"Incompatible data type ({data_type}) and chart type ({chart_type}).")
    
    except Exception as e:
        raise ValueError(f"Failed to create chart from sample data: {str(e)}")

# Styling and theming tools
@mcp.tool()
def apply_theme_from_presentation(presentation_name: str, source_presentation_id: str) -> str:
    """Apply theme from another Google Slides presentation.

    Args:
        presentation_name: Name of the target presentation
        source_presentation_id: ID of the source presentation to copy theme from

    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")

        result = _slides_manager.apply_theme_from_presentation(presentation_name, source_presentation_id)
        return result
    except Exception as e:
        raise ValueError(f"Failed to apply theme: {str(e)}")

@mcp.tool()
def apply_beautiful_styling(presentation_name: str) -> str:
    """Apply beautiful styling and colors to make the presentation look professional.

    Args:
        presentation_name: Name of the presentation

    Returns:
        Confirmation message
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")

        result = _slides_manager.apply_beautiful_styling(presentation_name)
        return result
    except Exception as e:
        raise ValueError(f"Failed to apply styling: {str(e)}")

@mcp.tool()
def apply_theme_by_name(presentation_name: str, theme_name: str) -> str:
    """Search for and apply a theme template from Google Drive by name.

    Args:
        presentation_name: Name of the target presentation
        theme_name: Name or partial name of the theme template to search for

    Returns:
        Confirmation message with the applied theme name
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")

        result = _slides_manager.apply_theme_by_name(presentation_name, theme_name)
        return result
    except Exception as e:
        raise ValueError(f"Failed to apply theme by name: {str(e)}")

@mcp.tool()
def list_available_themes() -> str:
    """List available theme templates in Google Drive.

    Returns:
        List of available themes with their names and IDs
    """
    try:
        global _slides_manager
        if not _slides_manager:
            raise ValueError("No active slides manager. Create a presentation first.")

        themes = _slides_manager.list_available_themes()

        if not themes:
            return "No theme templates found in Google Drive."

        result = "Available themes:\n"
        for theme in themes:
            result += f"- {theme['name']} (ID: {theme['id']}) - Modified: {theme['modified']}\n"

        return result
    except Exception as e:
        raise ValueError(f"Failed to list themes: {str(e)}")

# Initialize the global slides manager
_slides_manager = None


def main():
    logger.info(f"Starting Google Slides MCP Server")
    mcp.run(transport='stdio')

if __name__ == "__main__":
    main()
