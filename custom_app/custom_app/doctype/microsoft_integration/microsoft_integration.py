import json
import frappe
import requests
from datetime import datetime
from frappe.model.document import Document

class MicrosoftIntegration(Document):
    pass

@frappe.whitelist()
def get_active_users():
    users = frappe.get_all('User', filters={'enabled': 1}, fields=['full_name', 'email'])
    return users

@frappe.whitelist()
def get_access_token():
    client_id = 'client-id'
    client_secret = 'client_secret'
    tenant_id = 'tenant_id'
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials'
    }

    response = requests.post(url, headers=headers, data=data)
    if response.status_code == 200:
        return response.json()
    else:
        error_message = f"Failed to get access token: {response.status_code}, {response.text}"
        print(error_message)
        raise Exception(error_message)

token_response = get_access_token()
print(token_response)

def get_user_id(token, email):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'https://graph.microsoft.com/v1.0/users?$filter=mail eq \'{email}\''
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        users = response.json()
        if users['value']:
            return users['value'][0]['id']
        else:
            error_message = f"User with email {email} not found"
            print(error_message)
            raise Exception(error_message)
    else:
        error_message = f"Failed to get user ID: {response.status_code}, {response.text}"
        print(error_message)
        raise Exception(error_message)

def check_existing_meeting(token, user_id, subject, start_iso, end_iso):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events?$filter=subject eq '{subject}' and start/dateTime eq '{start_iso}' and end/dateTime eq '{end_iso}'"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        meetings = response.json().get('value', [])
        if meetings:
            return meetings[0]
        return None
    else:
        error_message = f"Failed to check existing meeting: {response.status_code}, {response.text}"
        frappe.log_error(error_message, 'Microsoft Graph API Error')
        frappe.throw(error_message)

def create_teams_meeting(token, subject, start_iso, end_iso, location, body, attendees_emails, user_id):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events"

    attendees_list = json.loads(attendees_emails)
    if not isinstance(attendees_list, list):
        raise ValueError("Attendees should be a list of email addresses.")

    meeting_data = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body
        },
        "start": {
            "dateTime": start_iso,
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_iso,
            "timeZone": "UTC"
        },
        "location": {
            "displayName": location
        },
        "attendees": [{"emailAddress": {"address": email}} for email in attendees_list]
    }

    response = requests.post(url, headers=headers, json=meeting_data)

    if response.status_code == 201:
        return {"status": "success", "message": response.json()}
    else:
        error_message = f"Failed to create meeting: {response.status_code}, {response.json()}"
        print(error_message)
        return {"status": "failure", "message": response.json()}


@frappe.whitelist()
def schedule_meeting(subject, body, start, end, location, attendees):
    token_response = get_access_token()

    if token_response and 'access_token' in token_response:
        token = token_response['access_token']

        email = 'ibrahimmujahid551@gmail.com'
        user_id = get_user_id(token, email)

        start_iso = datetime.strptime(start, "%d-%m-%Y %H:%M:%S").isoformat() + "Z"
        end_iso = datetime.strptime(end, "%d-%m-%Y %H:%M:%S").isoformat() + "Z"

        response = create_teams_meeting(token, subject, start_iso, end_iso, location, body, attendees, user_id)

        if response.get("error"):
            return {'status': 'failure', 'message': response.get("error")}
        else:
            return {'status': 'success', 'message': 'Meeting scheduled successfully'}
    else:
        error_message = 'Unable to retrieve access token'
        print(error_message)
        return {'status': 'failure', 'message': error_message}
    










def log_error_truncated(title, message):
    if len(message) > 140:
        message = message[:137] + '...'
    frappe.log_error(message, title)

# Example usage
log_error_truncated('Microsoft Graph API Error', 'Your error message here...')



