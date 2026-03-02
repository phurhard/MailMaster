import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from api.utils.google_apis import create_service


def init_gmail_service(client_file, api_name='gmail', api_version='v1', scopes=['https://mail.google.com']):
    return create_service(client_file, api_name, api_version, scopes)


def _extract_body(payload):
    """Recursively extract body content, prioritizing text/html."""
    html_content = ""
    text_content = ""

    def walk_parts(parts):
        nonlocal html_content, text_content
        for part in parts:
            mime = part.get('mimeType')
            data = part.get('body', {}).get('data')
            if mime == 'text/plain' and data:
                text_content = base64.urlsafe_b64decode(data).decode('utf-8')
            elif mime == 'text/html' and data:
                html_content = base64.urlsafe_b64decode(data).decode('utf-8')
            elif 'parts' in part:
                walk_parts(part['parts'])

    if 'parts' in payload:
        walk_parts(payload['parts'])
    elif 'body' in payload and 'data' in payload['body']:
        data = payload['body']['data']
        mime = payload.get('mimeType')
        if mime == 'text/plain':
            text_content = base64.urlsafe_b64decode(data).decode('utf-8')
        elif mime == 'text/html':
            html_content = base64.urlsafe_b64decode(data).decode('utf-8')

    return html_content or text_content or '<Body not available>'

def _extract_attachments(payload):
    """Recursively extract attachment metadata."""
    attachments = []
    if not payload:
        return []
    if 'parts' in payload:
        for part in payload['parts']:
            filename = part.get('filename')
            if filename:
                body = part.get('body', {})
                attachments.append({
                    'filename': filename,
                    'mimeType': part.get('mimeType'),
                    'id': body.get('attachmentId'),
                    'size': body.get('size', 0)
                })
            elif 'parts' in part:
                attachments.extend(_extract_attachments(part))
    return attachments

def list_labels(service):
    """List all labels for the user."""
    results = service.users().labels().list(userId='me').execute()
    return results.get('labels', [])

def ensure_label_exists(service, label_name):
    """Checks if a label exists, if not creates it. Returns label ID."""
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
        if label['name'] == label_name:
            return label['id']
    
    # Create label
    label_body = {
        'name': label_name,
        'labelListVisibility': 'labelShow',
        'messageListVisibility': 'show'
    }
    created_label = service.users().labels().create(userId='me', body=label_body).execute()
    return created_label['id']

def modify_email_labels(service, user_id, message_id, add_labels=None, remove_labels=None):
    """Apply/Remove labels from a message."""
    body = {}
    if add_labels:
        body['addLabelIds'] = add_labels
    if remove_labels:
        body['removeLabelIds'] = remove_labels
    
    if not body:
        return
        
    return service.users().messages().modify(
        userId=user_id,
        id=message_id,
        body=body
    ).execute()


def get_email_messages(service, user_id='me', label_ids=None, folder_name='INBOX', max_results=10):
    messages = []
    next_page_token = None

    if folder_name:
        label_results = service.users().labels().list(userId=user_id).execute()
        labels = label_results.get('labels', [])
        folder_label_id = next((label['id'] for label in labels if label['name'].lower() == folder_name.lower()), None)
        if folder_label_id:
            if label_ids:
                label_ids.append(folder_label_id)
            else:
                label_ids = [folder_label_id]
        else:
            raise ValueError(f"Folder '{folder_name}' not found")
    
    while True:
        result = service.users().messages().list(
            userId=user_id,
            labelIds=label_ids,
            maxResults=min(500, max_results - len(messages)) if max_results else 500,
            pageToken=next_page_token
        ).execute()

        messages.extend(result.get('messages', []))

        next_page_token = result.get('nextPageToken')

        if not next_page_token or (max_results and len(messages) >= max_results):
            break
    
    return messages[:max_results] if max_results else messages

def get_email_message_details(service, msg_id):
    message = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
    payload = message['payload']
    headers = payload.get('headers', [])

    subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), message.get('subject', 'No subject'))
    sender = next((header['value'] for header in headers if header['name'].lower() == 'from'), 'No sender')
    recipients = next((header['value'] for header in headers if header['name'].lower() == 'to'), 'No recipients')
    snippet = message.get('snippet', 'No snippet')
    has_attachments = any(part.get('filename') for part in payload.get('parts', []) if part.get('filename'))
    date = next((header['value'] for header in headers if header['name'].lower() == 'date'), 'No date')
    star = message.get('labelIds', []).count('STARRED') > 0
    label = ', '.join(message.get('labelIds', []))

    # Extended headers for categorization
    list_unsubscribe = next((header['value'] for header in headers if header['name'].lower() == 'list-unsubscribe'), None)
    x_mailer = next((header['value'] for header in headers if header['name'].lower() == 'x-mailer'), None)
    precedence = next((header['value'] for header in headers if header['name'].lower() == 'precedence'), None)

    body = _extract_body(payload)
    attachments = _extract_attachments(payload)

    return {
        'subject': subject,
        'sender': sender,
        'recipients': recipients,
        'body': body,
        'snippet': snippet,
        'has_attachments': len(attachments) > 0,
        'attachments': attachments,
        'date': date,
        'star': star,
        'label': label,
        'size_estimate': message.get('sizeEstimate', 0),
        'id': msg_id,
        'headers': {
            'list_unsubscribe': list_unsubscribe,
            'x_mailer': x_mailer,
            'precedence': precedence
        }
    }


def get_batch_email_details(service, msg_ids):
    if not msg_ids:
        return []
    
    results = {}
    def callback(request_id, response, exception):
        if exception is not None:
            import logging
            logging.getLogger(__name__).error(f"Error fetching message {request_id}: {exception}")
        else:
            results[request_id] = response
            
    # Process in smaller chunks to avoid Gmail's "Too many concurrent requests for user" limit
    import time
    CHUNK_SIZE = 20
    for i in range(0, len(msg_ids), CHUNK_SIZE):
        chunk = msg_ids[i:i+CHUNK_SIZE]
        batch = service.new_batch_http_request(callback=callback)
        for msg_id in chunk:
            batch.add(
                service.users().messages().get(userId='me', id=msg_id, format='full'),
                request_id=msg_id
            )
        batch.execute()
        
        # Add a small delay between batches if there are more emails to fetch
        if i + CHUNK_SIZE < len(msg_ids):
            time.sleep(0.5)
        
    processed_emails = []
    # maintain original order
    for msg_id in msg_ids:
        message = results.get(msg_id)
        if not message:
            continue
            
        payload = message.get('payload', {})
        headers = payload.get('headers', [])
        
        subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), message.get('subject', 'No subject'))
        sender = next((header['value'] for header in headers if header['name'].lower() == 'from'), 'No sender')
        recipients = next((header['value'] for header in headers if header['name'].lower() == 'to'), 'No recipients')
        snippet = message.get('snippet', 'No snippet')
        has_attachments = any(part.get('filename') for part in payload.get('parts', []) if part.get('filename'))
        date = next((header['value'] for header in headers if header['name'].lower() == 'date'), 'No date')
        star = message.get('labelIds', []).count('STARRED') > 0
        label_list = message.get('labelIds', [])
        label = ', '.join(label_list) if isinstance(label_list, list) else ''
        
        # Extended headers for categorization
        list_unsubscribe = next((header['value'] for header in headers if header['name'].lower() == 'list-unsubscribe'), None)
        x_mailer = next((header['value'] for header in headers if header['name'].lower() == 'x-mailer'), None)
        precedence = next((header['value'] for header in headers if header['name'].lower() == 'precedence'), None)

        body = _extract_body(payload)
        attachments = _extract_attachments(payload)
        
        processed_emails.append({
            'subject': subject,
            'sender': sender,
            'recipients': recipients,
            'body': body,
            'snippet': snippet,
            'has_attachments': len(attachments) > 0,
            'attachments': attachments,
            'date': date,
            'star': star,
            'label': label,
            'size_estimate': message.get('sizeEstimate', 0),
            'id': msg_id,
            'headers': {
                'list_unsubscribe': list_unsubscribe,
                'x_mailer': x_mailer,
                'precedence': precedence
            }
        })
        
    return processed_emails


def send_email(service, to, subject, body, body_type='plain', attachment_paths=None):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject

    if body_type.lower() not in ['plain', 'html']:
        raise ValueError('body_type must be eother "plain" or "html"')
    
    message.attach(MIMEText(body, body_type.lower()))

    if attachment_paths:
        for attachment_path in attachment_paths:
            if os.path.exists(attachment_path):
                filename = os.path.basename(attachment_path)

                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                encoders.encode_base64(part)

                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
                )

                message.attach(part)
            else:
                raise FileNotFoundError(f"File not found - {attachment_path}")
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

    sent_message = service.users().messages().send(
        userId='me',
        body={'raw': raw_message}
    ).execute()
    return sent_message


def get_attachment_data(service, user_id, message_id, attachment_id):
    """Fetch the raw attachment data."""
    attachment = service.users().messages().attachments().get(
        userId=user_id, messageId=message_id, id=attachment_id
    ).execute()
    data = attachment.get('data')
    if not data:
        return b""
    return base64.urlsafe_b64decode(data.encode('UTF-8'))


def download_attachments_parent(service, user_id, msg_id, target_dir):
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    for part in message['payload']['parts']:
        if part['filename']:
            att_id = part['body']['attachmentId']
            att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=att_id).execute()
            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('utf-8'))
            file_path = os.path.join(target_dir, part['filename'])
            print('Saving attachment to : ', file_path)
            with open(file_path, 'wb') as f:
                f.write(file_data)
            
def download_attachments_all(service, user_id, msg_id, target_dir):
    thread = service.users().threads().get(userId=user_id, id=msg_id).execute()
    for message in thread['messages']:
        for part in message['payload']['parts']:
            if part['filename']:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(userId=user_id, messageId=message['id'], id=att_id).execute()
                data = att['data']
                file_data = base64.urlsafe_b64decode(data.encode('utf-8'))
                file_path = os.path.join(target_dir, part['filenmae'])
                print('Saving attachment to: ', file_path)
                with open(file_path, 'wb') as f:
                    f.write(file_data)


def search_emails(service, query, user_id='me', max_results=5):
    messages = []
    next_page_token = None

    while True:
        result = service.users().messages().list(
            userId=user_id,
            q=query,
            maxResults=min(500, max_results - len(messages)) if max_results else 500,
            pageToken=next_page_token
        ).execute()

        messages.extend(result.get('messages', []))

        next_page_token = result.get('nextPageToken')

        if not next_page_token or (max_results and len(messages) >= max_results):
            break
    return messages[:max_results] if max_results else messages

def list_large_emails(service, min_size_mb=5, max_results=20):
    """
    Find emails larger than a certain size.
    query: larger:5M
    """
    query = f"larger:{min_size_mb}M"
    return search_emails(service, query, max_results=max_results)


def search_email_conversations(service, query, user_id='me', max_results=5):
    conversations = []
    next_page_token = None

    while True:
        result = service.users().threads().list(
            userId=user_id,
            q=query,
            maxResults=min(500, max_results - len(conversations)) if max_results else 500,
            pageToken=next_page_token
        ).execute()

        conversations.extend(result.get('threads', []))

        next_page_token = result.get('nextPageToken')

        if not next_page_token or (max_results and len(conversations) >= max_results):
            break
    
    return conversations[:max_results] if max_results else conversations


def create_label(service, name, label_list_visibility='labelShow', message_list_visibility='show'):
    label = {
        'name': name,
        'labelListVisibility': label_list_visibility,
        'messageListVisibility': message_list_visibility
    }
    created_label = service.users().labels().create(userId='me', body=label).execute()
    return created_label

def list_labels(service):
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    return labels


def get_label_details(service, label_id):
    return service.users().labels().get(userId='me', id=label_id).execute()


def modify_label(service, label_id, **updates):
    label = service.users().labels().get(userId='me', id=label_id).execute()
    for key, value in updates.items():
        label[key] = value
    updated_label = service.users().labels().update(userId='me', id=label_id, body=label).execute()
    return updated_label


def delete_label(service, label_id):
    service.users().labels().delete(userID='me', id=label_id).execute()


def map_label_name_to_id(service, label_name):
    labels = list_labels(service)
    label = next((label for label in labels if label['name'] == label_name), None)
    return label['id'] if label else None

# manage email labels
def modify_email_labels(service, user_id, message_id, add_labels=None, remove_labels=None):
    def batch_labels(labels, batch_size=100):
        return [labels[i:i + batch_size] for i in range(0, len(labels), batch_size)]

    if add_labels:
        for batch in batch_labels(add_labels):
            service.users().messages().modify(
                userId=user_id,
                id=message_id,
                body={'addLabelIds': batch}
            ).execute()
    
    if remove_labels:
        for batch in batch_labels(remove_labels):
            service.users().messages().modify(
                userId=user_id,
                id=message_id,
                body={'removeLabelIds': batch}
            ).execute()


def trash_email(service, user_id, message_id):
    service.users().messages().trash(userId=user_id, id=message_id).execute()


def batch_trash_emails(service, user_id, message_ids):
    batch = service.new_batch_http_request()
    for message_id in message_ids:
        batch.add(service.users().messages().trash(userId=user_id, id=message_id))
    batch.execute()


def permanantly_delete_email(service, user_id, message_id):
    service.users().messages().delete(userId=user_id, id=message_id).execute()


def untrash_email(service, user_id, message_id):
    service.users().messages().untrash(userId=user_id, id=message_id).execute()


def batch_untrash_emails(service, user_id, message_ids):
    batch = service.new_batch_http_request()
    for message_id in message_ids:
        batch.add(service.users().messages().untrash(userId=user_id, id=message_id))
    batch.execute()


def empty_trash(service):
    page_token = None
    total_deleted = 0

    while True:
        response = service.users().messages().list(
            userId='me',
            q='in:trash',
            pageToken=page_token,
            maxResults=500
        ).execute()

        messages = response.get('messages', [])
        if not messages:
            break

        batch = service.new_batch_http_request()

        for message in messages:
            batch.add(service.users().messages().delete(userId='me', id=message['id']))
        batch.execute()

        total_deleted += len(message)

        page_token = response.get('nextPageToken')
        if not page_token:
            break
    return total_deleted

# manage drafts
def list_draft_email_messages(service, user_id='me', max_results=10):
    drafts = []
    next_page_token = None

    while True:
        result = service.users().drafts().list(
            userID=user_id,
            maxResults=min(500, max_results - len(drafts)) if max_results else 500,
            pageToken=next_page_token
        ).execute()

        drafts.extend(result.get('drafts', []))
        next_page_token = result.get('nextPageToken')

        if not next_page_token or (max_results and len(drafts) >= max_results):
            break
    
    return drafts[:max_results] if max_results else drafts



def create_draft_email(service, to, subject, body, body_type='plain', attachment_paths=None):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject

    if body_type.lower() not in ['plain', 'html']:
        raise ValueError('body_type must be eother "plain" or "html"')
    
    message.attach(MIMEText(body, body_type.lower()))

    if attachment_paths:
        for attachment_path in attachment_paths:
            if os.path.exists(attachment_path):
                filename = os.path.basename(attachment_path)

                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                encoders.encode_base64(part)

                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
                )

                message.attach(part)
            else:
                raise FileNotFoundError(f"File not found - {attachment_path}")
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

    draft_message = service.users().drafts().create(
        userId='me',
        body={'raw': raw_message}
    ).execute()
    return draft_message


def get_draft_email_message_details(service, draft_id, format='full'):
    draft_detail = service.users().drafts().get(userId='me', id=draft_id, format=format).execute()
    message = draft_detail['message']
    payload = message['payload']
    headers = payload.get('headers', [])

    subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), "No subject")    
    sender = next((header['value'] for header in headers if header['name'] == 'From'), 'No sender')
    recipients = next((header['value'] for header in headers if header['name'] == 'To'), 'No recipients')
    snippet = message.get('snippet', 'No snippet')
    has_attachments = any(part.get('filename') for part in payload.get('parts', []) if part.get('filename'))
    date = next((header['value'] for header in headers if header['name'] == 'Date'), 'No date')
    star = message.get('labelsIds', []).count('STARRED') > 0
    label = ', '.join(message.get('labelIds', []))

    body = _extract_body(payload)

    return {
        'subject': subject,
        'sender': sender,
        'recipients': recipients,
        'body': body,
        'snippet': snippet,
        'has_attachments': has_attachments,
        'date': date,
        'star': star,
        'label': label,
        'size_estimate': message.get('sizeEstimate', 0),
        'id': msg_id
    }

def send_draft_email(service, draft_id):
    return service.users().drafts().send(userId='me', body={'id': draft_id}).execute()


def delete_draft_email(service, draft_id):
    service.users().drafts().delete(userId='me', id=draft_id).execute()


def get_message_and_replies(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id, format='minimal').execute()
    thread_id = message['threadId']
    thread = service.users().threads().get(userId='me', id=thread_id).execute()

    processed_messages = []

    for msg in thread['messages']:
        subject = next((header['value'] for header in msg['payload']['headers'] if header['name'].lower() == 'subject'), 'No Subject')
        from_header = next((header['value'] for header in msg['payload']['headers'] if header['name'].lower() == 'from'), 'Unknown Sender')
        date = next((header['value'] for header in msg['payload']['headers'] if header['name'].lower() == 'date'), 'Unknown Date')

        content = _extract_body(msg['payload'])

        processed_messages.append({
            'id': msg['id'],
            'subject': subject,
            'from': from_header,
            'date': date,
            'body': content
        })
    return processed_messages
