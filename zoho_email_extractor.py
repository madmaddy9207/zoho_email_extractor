import os
import json
import requests
import time
import logging
import sys
import re
from urllib.parse import urlencode, parse_qs, urlparse
from datetime import datetime
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
import pandas as pd
from collections import defaultdict
import email.utils

# Setup logging
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('zoho_extractor.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ZohoEmailExtractor:
    def __init__(self):
        self.client_id = os.getenv('ZOHO_CLIENT_ID')
        self.client_secret = os.getenv('ZOHO_CLIENT_SECRET')
        self.redirect_uri = os.getenv('ZOHO_REDIRECT_URI', 'http://localhost:5000/oauth/callback')
        self.access_token = None
        self.refresh_token = None
        self.token_expires_at = None
        self.base_url = "https://mail.zoho.in/api"
        self.account_id = None
        
        # Rate limiting settings based on Zoho API limits
        self.rate_limit_delay = 1.2  # Increased delay
        self.max_retries = 3
        self.requests_per_minute = 40  # More conservative limit
        self.request_timestamps = []
        
        # Pagination settings
        self.batch_size = 50  # Increased batch size for efficiency
        self.max_messages = 5000  # Maximum messages to process
        
        # Attachment settings
        self.download_attachments = True
        self.attachment_api_available = True  # Will be set to False if API doesn't support attachments
        self.max_attachment_size = 10 * 1024 * 1024  # 10MB limit
        self.allowed_extensions = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.csv', '.zip', '.rar', '.jpg', '.jpeg', '.png', '.gif'}
        
        # Create output directory
        self.output_dir = "zoho_email_extraction"
        self.attachments_dir = os.path.join(self.output_dir, "attachments")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.attachments_dir, exist_ok=True)
        
        if not all([self.client_id, self.client_secret]):
            raise ValueError("Please set ZOHO_CLIENT_ID and ZOHO_CLIENT_SECRET environment variables")

    def rate_limit_check(self):
        """Implement rate limiting to avoid API limits"""
        current_time = time.time()
        
        # Remove timestamps older than 1 minute
        self.request_timestamps = [ts for ts in self.request_timestamps if current_time - ts < 60]
        
        # If we're approaching the limit, wait
        if len(self.request_timestamps) >= self.requests_per_minute - 5:
            wait_time = 60 - (current_time - self.request_timestamps[0])
            if wait_time > 0:
                logger.info(f"Rate limit approaching, waiting {wait_time:.1f} seconds...")
                time.sleep(wait_time)
                self.request_timestamps = []
        
        # Add delay between requests
        time.sleep(self.rate_limit_delay)
        self.request_timestamps.append(current_time)

    def get_authorization_url(self):
        """Generate authorization URL for OAuth2 flow"""
        auth_url = "https://accounts.zoho.in/oauth/v2/auth"
        params = {
            'response_type': 'code',
            'client_id': self.client_id,
            'scope': 'ZohoMail.messages.READ,ZohoMail.folders.READ,ZohoMail.accounts.READ',
            'redirect_uri': self.redirect_uri,
            'access_type': 'offline',
            'prompt': 'consent'  # Force consent to ensure refresh token
        }
        return f"{auth_url}?{urlencode(params)}"

    def exchange_code_for_tokens(self, auth_code):
        """Exchange authorization code for access and refresh tokens"""
        token_url = "https://accounts.zoho.in/oauth/v2/token"
        data = {
            'grant_type': 'authorization_code',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'redirect_uri': self.redirect_uri,
            'code': auth_code
        }
        
        try:
            response = requests.post(token_url, data=data, timeout=30)
            logger.info(f"Token exchange response: {response.status_code}")
            
            if response.status_code == 200:
                tokens = response.json()
                self.access_token = tokens.get('access_token')
                self.refresh_token = tokens.get('refresh_token')
                
                # Calculate expiration time
                expires_in = tokens.get('expires_in', 3600)
                self.token_expires_at = time.time() + expires_in - 300  # 5 min buffer
                
                # Add metadata
                tokens['expires_at'] = self.token_expires_at
                tokens['retrieved_at'] = time.time()
                
                # Save tokens securely
                with open(os.path.join(self.output_dir, 'tokens.json'), 'w') as f:
                    json.dump(tokens, f, indent=2)
                
                logger.info("Tokens saved successfully")
                logger.info(f"Access token expires in {expires_in} seconds")
                logger.info(f"Refresh token available: {bool(self.refresh_token)}")
                return True
            else:
                logger.error(f"Token exchange failed: {response.status_code} - {response.text}")
                return False
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error during token exchange: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during token exchange: {e}")
            return False

    def load_tokens(self):
        """Load saved tokens from file"""
        token_file = os.path.join(self.output_dir, 'tokens.json')
        try:
            if not os.path.exists(token_file):
                logger.info("No token file found")
                return False
                
            with open(token_file, 'r') as f:
                content = f.read().strip()
                if not content:
                    logger.warning("Token file is empty")
                    return False
                    
                tokens = json.loads(content)
                self.access_token = tokens.get('access_token')
                self.refresh_token = tokens.get('refresh_token')
                self.token_expires_at = tokens.get('expires_at', 0)
                
                # Check if token is expired
                if self.token_expires_at and time.time() > self.token_expires_at:
                    logger.info("Saved token is expired, will need refresh")
                    if self.refresh_token:
                        return self.refresh_access_token()
                    else:
                        logger.warning("No refresh token available for expired access token")
                        return False
                
                if self.access_token:
                    logger.info("Tokens loaded successfully")
                    return True
                else:
                    logger.warning("No access token found in file")
                    return False
                    
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in token file: {e}")
            try:
                os.remove(token_file)
                logger.info("Corrupted token file deleted")
            except:
                pass
            return False
        except Exception as e:
            logger.error(f"Error loading tokens: {e}")
            return False

    def refresh_access_token(self):
        """Refresh the access token using refresh token"""
        if not self.refresh_token:
            logger.warning("No refresh token available")
            return False
            
        token_url = "https://accounts.zoho.in/oauth/v2/token"
        data = {
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.refresh_token
        }
        
        try:
            logger.info("Attempting to refresh access token...")
            response = requests.post(token_url, data=data, timeout=30)
            
            if response.status_code == 200:
                tokens = response.json()
                self.access_token = tokens.get('access_token')
                
                # Update expiration time
                expires_in = tokens.get('expires_in', 3600)
                self.token_expires_at = time.time() + expires_in - 300
                
                # Update saved tokens
                token_file = os.path.join(self.output_dir, 'tokens.json')
                try:
                    # Load existing tokens to preserve refresh token
                    with open(token_file, 'r') as f:
                        saved_tokens = json.load(f)
                    
                    # Update with new access token
                    saved_tokens['access_token'] = self.access_token
                    saved_tokens['expires_at'] = self.token_expires_at
                    saved_tokens['refreshed_at'] = time.time()
                    
                    # Update refresh token if provided
                    if tokens.get('refresh_token'):
                        saved_tokens['refresh_token'] = tokens['refresh_token']
                        self.refresh_token = tokens['refresh_token']
                    
                    with open(token_file, 'w') as f:
                        json.dump(saved_tokens, f, indent=2)
                    
                    logger.info("Access token refreshed successfully")
                    return True
                    
                except Exception as e:
                    logger.error(f"Error updating token file: {e}")
                    return False
                
            else:
                logger.error(f"Token refresh failed: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Exception during token refresh: {e}")
            return False

    def is_token_expired(self):
        """Check if access token is expired or will expire soon"""
        if not self.token_expires_at:
            return False
        return time.time() > self.token_expires_at

    def ensure_valid_token(self):
        """Ensure we have a valid access token"""
        if not self.access_token:
            raise ValueError("No access token available")
        
        if self.is_token_expired():
            logger.info("Token expired, refreshing...")
            if not self.refresh_access_token():
                raise ValueError("Failed to refresh expired token")

    def make_api_request(self, endpoint, method='GET', params=None, max_retries=None):
        """Make authenticated API request with retry logic and rate limiting"""
        if max_retries is None:
            max_retries = self.max_retries
        
        # Ensure we have a valid token
        self.ensure_valid_token()
        
        # Apply rate limiting
        self.rate_limit_check()
        
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json',
            'User-Agent': 'ZohoEmailExtractor/1.0'
        }
        
        url = f"{self.base_url}/{endpoint}"
        
        for attempt in range(max_retries + 1):
            try:
                logger.debug(f"API Request attempt {attempt + 1}: {method} {url}")
                
                if method == 'GET':
                    response = requests.get(url, headers=headers, params=params, timeout=30)
                else:
                    response = requests.request(method, url, headers=headers, params=params, timeout=30)
                
                if response.status_code == 401:
                    logger.info("Got 401, attempting to refresh token...")
                    if self.refresh_access_token():
                        headers['Authorization'] = f'Zoho-oauthtoken {self.access_token}'
                        if method == 'GET':
                            response = requests.get(url, headers=headers, params=params, timeout=30)
                        else:
                            response = requests.request(method, url, headers=headers, params=params, timeout=30)
                    else:
                        raise Exception("Failed to refresh token after 401 error")
                
                if response.status_code == 429:  # Rate limit exceeded
                    wait_time = min(2 ** attempt, 60)  # Exponential backoff, max 60 seconds
                    logger.warning(f"Rate limit hit, waiting {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue
                
                if response.status_code >= 500:  # Server error
                    if attempt < max_retries:
                        wait_time = min(2 ** attempt, 30)
                        logger.warning(f"Server error ({response.status_code}), retrying in {wait_time} seconds...")
                        time.sleep(wait_time)
                        continue
                
                # Log response for debugging
                if response.status_code != 200:
                    logger.warning(f"API request returned {response.status_code}: {response.text}")
                
                return response
                
            except requests.exceptions.Timeout:
                if attempt < max_retries:
                    logger.warning(f"Request timeout, retrying... (attempt {attempt + 1})")
                    time.sleep(2 ** attempt)
                    continue
                else:
                    raise Exception("Request timed out after multiple attempts")
            
            except requests.exceptions.RequestException as e:
                if attempt < max_retries:
                    logger.warning(f"Network error, retrying... (attempt {attempt + 1}): {e}")
                    time.sleep(2 ** attempt)
                    continue
                else:
                    raise Exception(f"Network error after multiple attempts: {e}")
        
        raise Exception(f"API request failed after {max_retries + 1} attempts")

    def get_account_info(self):
        """Get account information and set account_id"""
        try:
            logger.info("Fetching account information...")
            response = self.make_api_request('accounts')
            
            if response.status_code == 200:
                accounts = response.json()
                logger.debug(f"Accounts response: {accounts}")
                
                if accounts.get('data') and len(accounts['data']) > 0:
                    self.account_id = accounts['data'][0]['accountId']
                    account_name = accounts['data'][0].get('displayName', 'Unknown')
                    logger.info(f"Account ID retrieved: {self.account_id} ({account_name})")
                    return True
                else:
                    logger.error("No accounts found in response")
                    return False
            else:
                logger.error(f"Failed to get account info: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Exception getting account info: {e}")
            return False

    def get_folders(self):
        """Get all folders to find the inbox folder ID"""
        try:
            logger.info("Fetching folders...")
            response = self.make_api_request(f'accounts/{self.account_id}/folders')
            
            if response.status_code == 200:
                folders = response.json()
                logger.debug(f"Folders response: {folders}")
                
                # Look for inbox folder (usually has folderName "Inbox" or "INBOX")
                inbox_folder = None
                if folders.get('data'):
                    for folder in folders['data']:
                        folder_name = folder.get('folderName', '').lower()
                        if folder_name in ['inbox', 'inbox folder']:
                            inbox_folder = folder
                            break
                    
                    # If no exact match, try the first folder or look for system folder
                    if not inbox_folder:
                        for folder in folders['data']:
                            if folder.get('systemFolder') or folder.get('folderId') == '1':
                                inbox_folder = folder
                                break
                    
                    # Fallback to first folder
                    if not inbox_folder and folders['data']:
                        inbox_folder = folders['data'][0]
                
                if inbox_folder:
                    folder_id = inbox_folder.get('folderId')
                    folder_name = inbox_folder.get('folderName', 'Unknown')
                    logger.info(f"Using folder: {folder_name} (ID: {folder_id})")
                    return folder_id
                else:
                    logger.error("No suitable folder found")
                    return None
            else:
                logger.error(f"Failed to get folders: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"Exception getting folders: {e}")
            return None

    def get_messages_batch(self, folder_id, start_index=0, limit=25):
        """Get a batch of messages with pagination using the correct endpoint"""
        if not self.account_id or not folder_id:
            return [], 0
        
        params = {
            'start': str(start_index),
            'limit': str(limit),
            'folderId': str(folder_id)
        }
        
        try:
            # Use the correct endpoint for getting emails in a folder
            response = self.make_api_request(f'accounts/{self.account_id}/messages/view', params=params)
            
            if response.status_code == 200:
                data = response.json()
                messages = data.get('data', [])
                total_count = data.get('total', 0)
                
                # Debug: Check the structure of the first message
                if messages and len(messages) > 0:
                    logger.info(f"First message type: {type(messages[0])}")
                    logger.info(f"First message sample: {str(messages[0])[:200]}")
                
                logger.info(f"Fetched batch: {len(messages)} messages (starting from {start_index})")
                logger.info(f"Total messages available: {total_count}")
                
                # Debug: Check the structure of the first message
                if messages and len(messages) > 0:
                    logger.info(f"First message type: {type(messages[0])}")
                    logger.info(f"First message sample: {str(messages[0])[:200]}")
                    if isinstance(messages[0], dict):
                        logger.info(f"Message keys: {list(messages[0].keys())}")
                
                return messages, total_count
            else:
                logger.error(f"Failed to fetch messages batch: {response.status_code} - {response.text}")
                # Try alternative approach with search endpoint
                return self.get_messages_batch_search(start_index, limit)
                
        except Exception as e:
            logger.error(f"Exception fetching messages batch: {e}")
            return [], 0

    def get_messages_batch_search(self, start_index=0, limit=25):
        """Alternative method using search endpoint"""
        if not self.account_id:
            return [], 0
        
        params = {
            'start': str(start_index),
            'limit': str(limit)
        }
        
        try:
            # Use search endpoint as fallback
            response = self.make_api_request(f'accounts/{self.account_id}/messages/search', params=params)
            
            if response.status_code == 200:
                data = response.json()
                messages = data.get('data', [])
                total_count = data.get('total', 0)
                
                logger.info(f"Fetched batch via search: {len(messages)} messages (starting from {start_index})")
                logger.info(f"Total messages available: {total_count}")
                
                # Debug: Check the structure of the first message
                if messages and len(messages) > 0:
                    logger.info(f"Search - First message type: {type(messages[0])}")
                    logger.info(f"Search - First message sample: {str(messages[0])[:200]}")
                    if isinstance(messages[0], dict):
                        logger.info(f"Search - Message keys: {list(messages[0].keys())}")
                
                return messages, total_count
            else:
                logger.error(f"Failed to fetch messages via search: {response.status_code} - {response.text}")
                return [], 0
                
        except Exception as e:
            logger.error(f"Exception fetching messages via search: {e}")
            return [], 0

    def get_message_details(self, message_id):
        """Get full message details from message ID"""
        try:
            response = self.make_api_request(f'accounts/{self.account_id}/messages/{message_id}')
            
            if response.status_code == 200:
                message_data = response.json()
                if message_data.get('data'):
                    return self.extract_email_from_full_message(message_data['data'])
            
            return None
            
        except Exception as e:
            logger.error(f"Error fetching message details for {message_id}: {e}")
            return None
    
    def extract_email_from_full_message(self, message_data):
        """Extract email info from full message data structure"""
        try:
            sender_email = message_data.get('fromAddress', '').strip().lower()
            sender_name = message_data.get('sender', {}).get('name', '').strip()
            
            if not sender_name:
                sender_name = message_data.get('fromName', '').strip()
            
            if not sender_name and sender_email:
                if '<' in sender_email:
                    parsed = email.utils.parseaddr(sender_email)
                    if parsed[0]:
                        sender_name = parsed[0].strip('"').strip()
                        sender_email = parsed[1].lower()
                else:
                    sender_name = sender_email.split('@')[0].replace('.', ' ').replace('_', ' ').title()
            
            if sender_email and '@' in sender_email:
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                if re.match(email_pattern, sender_email):
                    email_info = {
                        'email': sender_email,
                        'name': sender_name or 'Unknown',
                        'subject': message_data.get('subject', '').strip(),
                        'received_time': message_data.get('receivedTime'),
                        'message_id': message_data.get('messageId') or message_data.get('id'),
                        'has_attachment': message_data.get('hasAttachment', False),
                        'attachments': []
                    }
                    
                    # Download attachments if enabled and message has attachments
                    if self.download_attachments and self.attachment_api_available and message_data.get('hasAttachment'):
                        try:
                            attachments = self.get_message_attachments(email_info['message_id'])
                            downloaded_attachments = []
                            
                            for attachment in attachments:
                                attachment_id = attachment.get('attachmentId')
                                filename = attachment.get('attachmentName', 'unknown')
                                
                                if attachment_id:
                                    file_path = self.download_attachment(
                                        email_info['message_id'], 
                                        attachment_id, 
                                        filename, 
                                        sender_email
                                    )
                                    if file_path:
                                        downloaded_attachments.append({
                                            'filename': filename,
                                            'path': file_path,
                                            'size': attachment.get('size', 0)
                                        })
                            
                            email_info['attachments'] = downloaded_attachments
                            
                        except Exception as e:
                            logger.debug(f"Attachment processing failed for {sender_email}: {e}")
                            self.attachment_api_available = False
                    
                    return email_info
            
            return None
            
        except Exception as e:
            logger.error(f"Error extracting from full message: {e}")
            return None

    def get_message_attachments(self, message_id):
        """Get attachments for a specific message"""
        try:
            # Try different possible endpoints for attachments
            endpoints = [
                f'accounts/{self.account_id}/messages/{message_id}/attachments',
                f'accounts/{self.account_id}/messages/{message_id}/attachment',
                f'accounts/{self.account_id}/folders/*/messages/{message_id}/attachments'
            ]
            
            for endpoint in endpoints:
                try:
                    response = self.make_api_request(endpoint)
                    if response.status_code == 200:
                        data = response.json()
                        return data.get('data', [])
                except:
                    continue
            
            logger.debug(f"No attachments endpoint found for message {message_id}")
            return []
                
        except Exception as e:
            logger.error(f"Error getting attachments for message {message_id}: {e}")
            return []
    
    def download_attachment(self, message_id, attachment_id, filename, sender_email):
        """Download a specific attachment"""
        try:
            # Create sender-specific directory
            sender_dir = os.path.join(self.attachments_dir, sender_email.replace('@', '_at_').replace('.', '_'))
            os.makedirs(sender_dir, exist_ok=True)
            
            # Clean filename
            safe_filename = "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
            if not safe_filename:
                safe_filename = f"attachment_{attachment_id}"
            
            file_path = os.path.join(sender_dir, safe_filename)
            
            # Check if file already exists
            if os.path.exists(file_path):
                logger.debug(f"Attachment already exists: {safe_filename}")
                return file_path
            
            # Check file extension
            file_ext = os.path.splitext(safe_filename)[1].lower()
            if file_ext not in self.allowed_extensions:
                logger.debug(f"Skipping attachment with disallowed extension: {safe_filename}")
                return None
            
            # Try different endpoints for downloading
            download_endpoints = [
                f'accounts/{self.account_id}/messages/{message_id}/attachments/{attachment_id}',
                f'accounts/{self.account_id}/messages/{message_id}/attachment/{attachment_id}',
                f'accounts/{self.account_id}/messages/{message_id}/attachments/{attachment_id}/content'
            ]
            
            for endpoint in download_endpoints:
                try:
                    response = self.make_api_request(endpoint)
                    if response.status_code == 200:
                        # Check file size
                        content_length = response.headers.get('content-length')
                        if content_length and int(content_length) > self.max_attachment_size:
                            logger.warning(f"Attachment too large, skipping: {safe_filename} ({content_length} bytes)")
                            return None
                        
                        # Save file
                        with open(file_path, 'wb') as f:
                            f.write(response.content)
                        
                        logger.info(f"Downloaded attachment: {safe_filename} from {sender_email}")
                        return file_path
                except:
                    continue
            
            logger.debug(f"Could not download attachment {attachment_id} - no working endpoint found")
            return None
                
        except Exception as e:
            logger.error(f"Error downloading attachment {attachment_id}: {e}")
            return None

    def extract_email_info(self, message):
        """Extract email and name information from a message"""
        try:
            # Handle case where message might be a string (message ID) instead of dict
            if isinstance(message, str):
                logger.info(f"Processing message ID: {message}")
                # Need to fetch full message details using the message ID
                return self.get_message_details(message)
            
            # Check if message is a dictionary
            if not isinstance(message, dict):
                logger.error(f"Unexpected message type: {type(message)}, value: {message}")
                return None
            
            # Extract sender information
            sender_email = message.get('fromAddress', '').strip().lower()
            
            # Handle sender field - it might be a string or dict
            sender_info = message.get('sender', {})
            # logger.debug(f"Sender info type: {type(sender_info)}, value: {sender_info}")
            
            if isinstance(sender_info, dict):
                sender_name = sender_info.get('name', '').strip()
            elif isinstance(sender_info, str):
                sender_name = sender_info.strip()
            else:
                sender_name = ''
            
            # If no name in sender object, try fromName
            if not sender_name:
                sender_name = message.get('fromName', '').strip()
            
            # If still no name, try to extract from email
            if not sender_name and sender_email:
                if '<' in sender_email:
                    # Format: "Name <email@domain.com>"
                    parsed = email.utils.parseaddr(sender_email)
                    if parsed[0]:
                        sender_name = parsed[0].strip('"').strip()
                        sender_email = parsed[1].lower()
                else:
                    # Use part before @ as name if no name provided
                    sender_name = sender_email.split('@')[0].replace('.', ' ').replace('_', ' ').title()
            
            # Clean and validate email
            if sender_email and '@' in sender_email:
                # Basic email validation
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                if re.match(email_pattern, sender_email):
                    email_info = {
                        'email': sender_email,
                        'name': sender_name or 'Unknown',
                        'subject': message.get('subject', '').strip(),
                        'received_time': message.get('receivedTime'),
                        'message_id': message.get('messageId') or message.get('id'),
                        'has_attachment': message.get('hasAttachment', False),
                        'attachments': []
                    }
                    
                    # Download attachments if enabled and message has attachments
                    if self.download_attachments and self.attachment_api_available and message.get('hasAttachment'):
                        try:
                            attachments = self.get_message_attachments(email_info['message_id'])
                            if not attachments and self.attachment_api_available:
                                # If we consistently can't get attachments, disable the feature
                                logger.warning("Attachment API appears to be unavailable, disabling attachment downloads")
                                self.attachment_api_available = False
                            
                            downloaded_attachments = []
                            for attachment in attachments:
                                attachment_id = attachment.get('attachmentId')
                                filename = attachment.get('attachmentName', 'unknown')
                                
                                if attachment_id:
                                    file_path = self.download_attachment(
                                        email_info['message_id'], 
                                        attachment_id, 
                                        filename, 
                                        sender_email
                                    )
                                    if file_path:
                                        downloaded_attachments.append({
                                            'filename': filename,
                                            'path': file_path,
                                            'size': attachment.get('size', 0)
                                        })
                            
                            email_info['attachments'] = downloaded_attachments
                            
                        except Exception as e:
                            logger.debug(f"Attachment processing failed for {sender_email}: {e}")
                            self.attachment_api_available = False
                    
                    return email_info
            
            return None
            
        except Exception as e:
            logger.error(f"Error extracting email info from message (type: {type(message)}): {e}")
            return None

    def extract_all_emails(self):
        """Extract all email addresses and names from inbox"""
        logger.info("Starting email extraction process...")
        
        # Get account info
        if not self.get_account_info():
            logger.error("Failed to get account information")
            return []
        
        # Get folder information
        folder_id = self.get_folders()
        if not folder_id:
            logger.warning("Could not get folder ID, trying without it...")
            folder_id = None
        
        all_emails = {}  # Use dict to automatically deduplicate by email
        processed_count = 0
        start_index = 0
        
        progress_file = os.path.join(self.output_dir, 'extraction_progress.json')
        
        try:
            while processed_count < self.max_messages:
                logger.info(f"Fetching batch starting at index {start_index}...")
                
                if folder_id:
                    messages, total_count = self.get_messages_batch(folder_id, start_index, self.batch_size)
                else:
                    messages, total_count = self.get_messages_batch_search(start_index, self.batch_size)
                
                if not messages:
                    logger.info("No more messages to process")
                    break
                
                batch_processed = 0
                for message in messages:
                    try:
                        email_info = self.extract_email_info(message)
                        if email_info and email_info['email']:
                            email_key = email_info['email']
                            
                            # If this email already exists, update with more recent info if applicable
                            if email_key in all_emails:
                                existing = all_emails[email_key]
                                # Keep the entry with more complete name information
                                if len(email_info['name']) > len(existing['name']) and email_info['name'] != 'Unknown':
                                    existing['name'] = email_info['name']
                                if email_info.get('subject') and len(email_info['subject']) > len(existing.get('subject', '')):
                                    existing['subject'] = email_info['subject']
                                # Update message count and timestamps
                                existing['message_count'] = existing.get('message_count', 1) + 1
                                existing['last_seen'] = email_info['received_time']
                                # Keep earliest first_seen
                                if not existing.get('first_seen') or email_info['received_time'] < existing['first_seen']:
                                    existing['first_seen'] = email_info['received_time']
                            else:
                                email_info['message_count'] = 1
                                email_info['first_seen'] = email_info['received_time']
                                email_info['last_seen'] = email_info['received_time']
                                all_emails[email_key] = email_info
                            
                            batch_processed += 1
                        
                    except Exception as e:
                        logger.error(f"Error processing individual message: {e}")
                        continue
                
                processed_count += len(messages)
                start_index += self.batch_size
                
                logger.info(f"Batch complete: {batch_processed} valid emails found")
                logger.info(f"Progress: {processed_count} messages processed, {len(all_emails)} unique emails found")
                
                # Show progress percentage if we have total count
                if total_count > 0:
                    progress_pct = min(100, (processed_count / total_count) * 100)
                    logger.info(f"Progress: {progress_pct:.1f}% complete")
                
                # Save progress periodically
                if processed_count % (self.batch_size * 5) == 0:
                    self.save_progress(all_emails, processed_count, progress_file)
                
                # Check if we've processed all available messages
                if len(messages) < self.batch_size:
                    logger.info("Reached end of available messages")
                    break
                
                # Small delay between batches
                time.sleep(1.0)
        
        except KeyboardInterrupt:
            logger.info("Extraction interrupted by user")
            logger.info(f"Saving progress... Found {len(all_emails)} unique emails so far")
            
        except Exception as e:
            logger.error(f"Unexpected error during extraction: {e}")
            
        finally:
            # Clean up progress file
            if os.path.exists(progress_file):
                try:
                    os.remove(progress_file)
                except:
                    pass
        
        email_list = list(all_emails.values())
        logger.info(f"Extraction complete! Found {len(email_list)} unique email addresses")
        
        return email_list

    def save_progress(self, emails_dict, processed_count, progress_file):
        """Save extraction progress to file"""
        try:
            progress_data = {
                'processed_count': processed_count,
                'unique_emails': len(emails_dict),
                'timestamp': datetime.now().isoformat(),
                'emails': list(emails_dict.values())
            }
            with open(progress_file, 'w', encoding='utf-8') as f:
                json.dump(progress_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Error saving progress: {e}")

    def save_to_excel(self, email_data):
        """Save extracted email data to Excel file"""
        try:
            if not email_data:
                logger.warning("No email data to save")
                return None
            
            # Prepare data for Excel
            excel_data = []
            for email_info in email_data:
                received_time = email_info.get('received_time')
                first_seen = email_info.get('first_seen', received_time)
                last_seen = email_info.get('last_seen', received_time)
                
                # Convert timestamps to readable dates
                received_date = ''
                first_seen_date = ''
                last_seen_date = ''
                
                try:
                    if received_time:
                        received_date = datetime.fromtimestamp(received_time/1000).strftime('%Y-%m-%d %H:%M:%S')
                    if first_seen:
                        first_seen_date = datetime.fromtimestamp(first_seen/1000).strftime('%Y-%m-%d %H:%M:%S')
                    if last_seen:
                        last_seen_date = datetime.fromtimestamp(last_seen/1000).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    pass
                
                # Count attachments
                attachment_count = len(email_info.get('attachments', []))
                attachment_files = ', '.join([att['filename'] for att in email_info.get('attachments', [])]) if attachment_count > 0 else ''
                
                excel_data.append({
                    'Email Address': email_info['email'],
                    'Name': email_info['name'],
                    'Message Count': email_info.get('message_count', 1),
                    'First Seen': first_seen_date,
                    'Last Seen': last_seen_date,
                    'Latest Subject': email_info.get('subject', '')[:100],  # Truncate long subjects
                    'Domain': email_info['email'].split('@')[1] if '@' in email_info['email'] else '',
                    'Has Attachments': email_info.get('has_attachment', False),
                    'Attachment Count': attachment_count,
                    'Attachment Files': attachment_files[:200]  # Truncate long lists
                })
            
            # Create DataFrame and save to Excel
            df = pd.DataFrame(excel_data)
            
            # Sort by message count (most frequent first)
            df = df.sort_values('Message Count', ascending=False)
            
            # Generate filename - use consistent name to avoid duplicates
            excel_file = os.path.join(self.output_dir, 'zoho_email_contacts_latest.xlsx')
            
            # If file exists, create backup
            if os.path.exists(excel_file):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = os.path.join(self.output_dir, f'zoho_email_contacts_backup_{timestamp}.xlsx')
                try:
                    os.rename(excel_file, backup_file)
                    logger.info(f"Previous file backed up as: {backup_file}")
                except Exception as e:
                    logger.warning(f"Could not backup existing Excel file: {e}")
                    # Try alternative filename if backup fails
                    excel_file = os.path.join(self.output_dir, f'zoho_email_contacts_{timestamp}.xlsx')
            
            # Create Excel writer with formatting
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                # Main data sheet
                df.to_excel(writer, sheet_name='Email Contacts', index=False)
                
                # Summary sheet
                summary_data = {
                    'Metric': [
                        'Total Unique Email Addresses',
                        'Total Messages Processed',
                        'Most Frequent Sender',
                        'Extraction Date',
                        'Unique Domains'
                    ],
                    'Value': [
                        len(df),
                        df['Message Count'].sum(),
                        df.iloc[0]['Email Address'] if len(df) > 0 else 'N/A',
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        df['Domain'].nunique()
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Domain analysis sheet
                domain_analysis = df.groupby('Domain').agg({
                    'Email Address': 'count',
                    'Message Count': 'sum'
                }).rename(columns={
                    'Email Address': 'Unique Emails',
                    'Message Count': 'Total Messages'
                }).sort_values('Total Messages', ascending=False)
                domain_analysis.to_excel(writer, sheet_name='Domain Analysis')
            
            logger.info(f"Excel file saved: {excel_file}")
            logger.info(f"Total unique emails: {len(df)}")
            logger.info(f"Total messages: {df['Message Count'].sum()}")
            
            return excel_file
            
        except PermissionError as e:
            logger.error(f"Permission denied saving Excel file (file may be open): {e}")
            # Try with timestamp filename
            try:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_file = os.path.join(self.output_dir, f'zoho_email_contacts_{timestamp}.xlsx')
                with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Email Contacts', index=False)
                logger.info(f"Excel file saved with timestamp: {excel_file}")
                return excel_file
            except Exception as e2:
                logger.error(f"Failed to save Excel file even with timestamp: {e2}")
                return None
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")
            return None

    def save_to_json(self, email_data):
        """Save extracted email data to JSON file"""
        try:
            if not email_data:
                logger.warning("No email data to save")
                return None
            
            # Generate filename - use consistent name to avoid duplicates
            json_file = os.path.join(self.output_dir, 'zoho_email_contacts_latest.json')
            
            # If file exists, create backup
            if os.path.exists(json_file):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = os.path.join(self.output_dir, f'zoho_email_contacts_backup_{timestamp}.json')
                try:
                    os.rename(json_file, backup_file)
                    logger.info(f"Previous JSON file backed up as: {backup_file}")
                except:
                    pass
            
            # Prepare data with readable timestamps
            json_data = []
            for email_info in email_data:
                processed_info = email_info.copy()
                
                # Convert timestamps to readable format
                for time_field in ['received_time', 'first_seen', 'last_seen']:
                    if processed_info.get(time_field):
                        try:
                            processed_info[f'{time_field}_readable'] = datetime.fromtimestamp(
                                processed_info[time_field]/1000
                            ).strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            processed_info[f'{time_field}_readable'] = 'Invalid Date'
                
                json_data.append(processed_info)
            
            # Sort by message count
            json_data.sort(key=lambda x: x.get('message_count', 0), reverse=True)
            
            # Save to JSON file
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'extraction_date': datetime.now().isoformat(),
                    'total_unique_emails': len(json_data),
                    'total_messages': sum(item.get('message_count', 0) for item in json_data),
                    'contacts': json_data
                }, f, indent=2, ensure_ascii=False)
            
            logger.info(f"JSON file saved: {json_file}")
            return json_file
            
        except Exception as e:
            logger.error(f"Error saving to JSON: {e}")
            return None

    def save_to_csv(self, email_data):
        """Save extracted email data to CSV file"""
        try:
            if not email_data:
                logger.warning("No email data to save")
                return None
            
            # Generate filename - use consistent name to avoid duplicates
            csv_file = os.path.join(self.output_dir, 'zoho_email_contacts_latest.csv')
            
            # If file exists, create backup
            if os.path.exists(csv_file):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = os.path.join(self.output_dir, f'zoho_email_contacts_backup_{timestamp}.csv')
                try:
                    os.rename(csv_file, backup_file)
                    logger.info(f"Previous CSV file backed up as: {backup_file}")
                except:
                    pass
            
            # Prepare data for CSV
            csv_data = []
            for email_info in email_data:
                received_time = email_info.get('received_time')
                first_seen = email_info.get('first_seen', received_time)
                last_seen = email_info.get('last_seen', received_time)
                
                # Convert timestamps to readable dates
                received_date = ''
                first_seen_date = ''
                last_seen_date = ''
                
                try:
                    if received_time:
                        received_date = datetime.fromtimestamp(received_time/1000).strftime('%Y-%m-%d %H:%M:%S')
                    if first_seen:
                        first_seen_date = datetime.fromtimestamp(first_seen/1000).strftime('%Y-%m-%d %H:%M:%S')
                    if last_seen:
                        last_seen_date = datetime.fromtimestamp(last_seen/1000).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    pass
                
                # Count attachments
                attachment_count = len(email_info.get('attachments', []))
                attachment_files = ', '.join([att['filename'] for att in email_info.get('attachments', [])]) if attachment_count > 0 else ''
                
                csv_data.append({
                    'email': email_info['email'],
                    'name': email_info['name'],
                    'message_count': email_info.get('message_count', 1),
                    'first_seen': first_seen_date,
                    'last_seen': last_seen_date,
                    'latest_subject': email_info.get('subject', ''),
                    'domain': email_info['email'].split('@')[1] if '@' in email_info['email'] else '',
                    'has_attachments': email_info.get('has_attachment', False),
                    'attachment_count': attachment_count,
                    'attachment_files': attachment_files
                })
            
            # Create DataFrame and save to CSV
            df = pd.DataFrame(csv_data)
            df = df.sort_values('message_count', ascending=False)
            df.to_csv(csv_file, index=False, encoding='utf-8')
            
            logger.info(f"CSV file saved: {csv_file}")
            return csv_file
            
        except Exception as e:
            logger.error(f"Error saving to CSV: {e}")
            return None

class OAuthHandler(BaseHTTPRequestHandler):
    """HTTP handler for OAuth callback"""
    
    def do_GET(self):
        """Handle GET request for OAuth callback"""
        try:
            # Parse the callback URL
            parsed_url = urlparse(self.path)
            query_params = parse_qs(parsed_url.query)
            
            if 'code' in query_params:
                # Store the authorization code
                self.server.auth_code = query_params['code'][0]
                
                # Send success response
                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(b'''
                <html>
                <head><title>Authorization Successful</title></head>
                <body>
                <h2>Authorization Successful!</h2>
                <p>You can close this window and return to the application.</p>
                <script>setTimeout(function(){window.close();}, 3000);</script>
                </body>
                </html>
                ''')
            elif 'error' in query_params:
                # Handle authorization error
                error = query_params.get('error', ['unknown'])[0]
                error_description = query_params.get('error_description', [''])[0]
                
                self.send_response(400)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(f'''
                <html>
                <head><title>Authorization Failed</title></head>
                <body>
                <h2>Authorization Failed</h2>
                <p>Error: {error}</p>
                <p>Description: {error_description}</p>
                <p>Please try again.</p>
                </body>
                </html>
                '''.encode())
                
        except Exception as e:
            logger.error(f"Error in OAuth handler: {e}")
            self.send_response(500)
            self.end_headers()
    
    def log_message(self, format, *args):
        """Suppress default HTTP server logging"""
        return

def start_oauth_server(port=5000):
    """Start local HTTP server for OAuth callback"""
    server = HTTPServer(('localhost', port), OAuthHandler)
    server.auth_code = None
    server.timeout = 300  # 5 minutes timeout
    
    def run_server():
        try:
            server.handle_request()
        except Exception as e:
            logger.error(f"OAuth server error: {e}")
    
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    
    return server

def print_banner():
    banner = r"""
███████╗ ██████╗ ██╗  ██╗ ██████╗     ███╗   ███╗ █████╗ ██╗██╗     
╚══███╔╝██╔═══██╗██║  ██║██╔═══██╗    ████╗ ████║██╔══██╗██║██║     
  ███╔╝ ██║   ██║███████║██║   ██║    ██╔████╔██║███████║██║██║     
 ███╔╝  ██║   ██║██╔══██║██║   ██║    ██║╚██╔╝██║██╔══██║██║██║     
███████╗╚██████╔╝██║  ██║╚██████╔╝    ██║ ╚═╝ ██║██║  ██║██║███████╗
╚══════╝ ╚═════╝ ╚═╝  ╚═╝ ╚═════╝     ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝
                                                                            
███████╗██╗  ██╗████████╗██████╗  █████╗  ██████╗████████╗ ██████╗ ██████╗  
██╔════╝╚██╗██╔╝╚══██╔══╝██╔══██╗██╔══██╗██╔════╝╚══██╔══╝██╔═══██╗██╔══██╗ 
█████╗   ╚███╔╝    ██║   ██████╔╝███████║██║        ██║   ██║   ██║██████╔╝ 
██╔══╝   ██╔██╗    ██║   ██╔══██╗██╔══██║██║        ██║   ██║   ██║██╔══██╗ 
███████╗██╔╝ ██╗   ██║   ██║  ██║██║  ██║╚██████╗   ██║   ╚██████╔╝██║  ██║ 
╚══════╝╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝ 
                                                                            
                 ───── Zoho Email Extractor v1 ─────
                    by SYSDEVCODE | Created by Abin P
"""
    print(banner)


def main():
    """Main function to run the email extractor"""
    try:
        # Initialize the extractor
        extractor = ZohoEmailExtractor()
        
        # Try to load existing tokens
        if not extractor.load_tokens():
            logger.info("No valid tokens found. Starting OAuth flow...")
            
            # Start OAuth server
            oauth_server = start_oauth_server()
            
            # Get authorization URL and open in browser
            auth_url = extractor.get_authorization_url()
            logger.info(f"Opening browser for authorization: {auth_url}")
            
            try:
                webbrowser.open(auth_url)
            except Exception as e:
                logger.warning(f"Could not open browser automatically: {e}")
                logger.info(f"Please manually open this URL in your browser: {auth_url}")
            
            # Wait for authorization code
            logger.info("Waiting for authorization callback...")
            timeout_counter = 0
            while oauth_server.auth_code is None and timeout_counter < 300:  # 5 minutes
                time.sleep(1)
                timeout_counter += 1
            
            if oauth_server.auth_code:
                logger.info("Authorization code received!")
                
                # Exchange code for tokens
                if extractor.exchange_code_for_tokens(oauth_server.auth_code):
                    logger.info("OAuth flow completed successfully!")
                else:
                    logger.error("Failed to exchange authorization code for tokens")
                    return
            else:
                logger.error("Authorization timeout. Please try again.")
                return
        
        # Now extract emails
        logger.info("Starting email extraction...")
        email_data = extractor.extract_all_emails()
        
        if email_data:
            logger.info(f"Successfully extracted {len(email_data)} unique email addresses")
            
            # Save in multiple formats
            excel_file = extractor.save_to_excel(email_data)
            json_file = extractor.save_to_json(email_data)
            csv_file = extractor.save_to_csv(email_data)
            
            # Print summary
            print("\n" + "="*60)
            print("EXTRACTION COMPLETE!")
            print("="*60)
            print(f"Unique email addresses found: {len(email_data)}")
            print(f"Total messages processed: {sum(item.get('message_count', 1) for item in email_data)}")
            
            # Calculate attachment statistics
            total_attachments = sum(len(item.get('attachments', [])) for item in email_data)
            emails_with_attachments = sum(1 for item in email_data if item.get('has_attachment', False))
            print(f"Emails with attachments: {emails_with_attachments}")
            print(f"Total attachments downloaded: {total_attachments}")
            if total_attachments > 0:
                print(f"Attachments saved to: {extractor.attachments_dir}")
            
            if excel_file:
                print(f"Excel file: {excel_file}")
            if json_file:
                print(f"JSON file: {json_file}")
            if csv_file:
                print(f"CSV file: {csv_file}")
            
            # Show top 10 most frequent senders
            email_data.sort(key=lambda x: x.get('message_count', 0), reverse=True)
            print(f"\nTop 10 most frequent senders:")
            for i, email_info in enumerate(email_data[:10], 1):
                attachment_info = f" [{len(email_info.get('attachments', []))} attachments]" if email_info.get('attachments') else ""
                print(f"{i:2d}. {email_info['name']} <{email_info['email']}> ({email_info.get('message_count', 1)} messages){attachment_info}")
            
            # Show senders with most attachments if any
            total_attachments = sum(len(item.get('attachments', [])) for item in email_data)
            if total_attachments > 0:
                attachment_senders = [(item, len(item.get('attachments', []))) for item in email_data if item.get('attachments')]
                attachment_senders.sort(key=lambda x: x[1], reverse=True)
                print(f"\nTop senders with attachments:")
                for i, (email_info, att_count) in enumerate(attachment_senders[:5], 1):
                    print(f"{i:2d}. {email_info['name']} <{email_info['email']}> ({att_count} attachments)")
            
            print("\n" + "="*60)

            print("Thank you for using Zoho Email Contact Extractor! By @AbinP from SYSDEVCODE")
            print("if you have any issues, please report them on the GitHub repository.")   
            print("if you like this tool, please consider giving it a star on GitHub!")
            print("if you have any suggestions or improvements, feel free to reach out! its an open-source project and contributions are welcome!")   
            print("connect with me on LinkedIn: https://www.linkedin.com/in/abinp-/")
            
        else:
            logger.warning("No email data extracted")
    
    except KeyboardInterrupt:
        logger.info("\nExtraction interrupted by user")
    except Exception as e:
        logger.error(f"Unexpected error in main: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print_banner()
    print("[INFO] Starting Zoho Email Extractor...")
    print("Zoho Email Contact Extractor By @AbinP from SYSDEVCODE")
    print("=" * 40)
    print("Make sure you have set the following environment variables:")
    print("- ZOHO_CLIENT_ID")
    print("- ZOHO_CLIENT_SECRET")
    print("- ZOHO_REDIRECT_URI (optional, defaults to http://localhost:5000/oauth/callback)")
    print("\nStarting extraction process...\n")
    
    main()
