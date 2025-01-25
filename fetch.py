from simplegmail import Gmail
from simplegmail.query import construct_query
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import re
import os
import pickle
import google.generativeai as genai
import json
import requests

# Google Gemini API Configuration
GENAI_API_KEY = "AIzaSyB6vX8Lddi4qMJlwnoiBxXEoII7gcwfw84"
genai.configure(api_key=GENAI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-pro-latest")

# API Endpoint
WEBSITE_API_URL = "http://localhost:3000/api/events"

results = []

def send_data_to_website(event_details):
    """
    Send event details to the API endpoint.
    """
    try:
        response = requests.post(WEBSITE_API_URL, json=event_details)
        if response.status_code == 200:
            print(f"Data sent successfully: {response.json()}")
        else:
            print(f"Failed to send data: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error while sending data: {e}")

def authenticate_calendar_api():
    """
    Authenticate and return Google Calendar API service instance.
    """
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('calendar', 'v3', credentials=creds)

def parse_travel_date(travel_date_str):
    """
    Parse the travel date from the given string using multiple formats.
    """
    formats = ['%a, %b %d, %Y', '%b %d, %Y', '%Y-%m-%d']
    for fmt in formats:
        try:
            return datetime.datetime.strptime(travel_date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date format: {travel_date_str}")

def delete_event(service, booking_reference, start_date, end_date):
    """
    Delete any existing events with the specified booking reference and dates.
    """
    try:
        service.events().delete(calendarId='primary', eventId=event_id).execute()
        events = service.events().list(
            calendarId='primary', q=booking_reference, singleEvents=True
        ).execute().get('items', [])

        for event in events:
            event_start = event['start'].get('dateTime') or event['start'].get('date')
            event_end = event['end'].get('dateTime') or event['end'].get('date')

            if start_date in event_start and end_date in event_end:
                service.events().delete(calendarId='primary', eventId=event['id']).execute()
                print(f"Deleted event: {event['summary']}")

    except Exception as e:
        print(f"Error deleting events: {e}")

def create_event(service, event_details):
    """
    Create a Google Calendar event from the extracted details.
    """
    try:
        # Parse travel date
        travel_date = parse_travel_date(event_details['Travel Date'])
        start_datetime = travel_date.replace(hour=9, minute=0).isoformat()
        end_datetime = travel_date.replace(hour=10, minute=0).isoformat()

        # Delete existing events before creating
        delete_event(service, event_details.get('Booking Reference', 'Unknown'), start_datetime, end_datetime)

        # Event body
        event = {
            'summary': event_details.get('Title', 'New Booking'),
            'description': f"""
                Booking Reference: {event_details.get('Booking Reference', 'Unknown')}
                Location: {event_details.get('Location')}
                Travel Date: {event_details['Travel Date']}
                Lead Traveler Name: {event_details.get('Lead Traveler Name')}
                Hotel Pickup: {event_details.get('Hotel Pickup')}
                Status: {event_details.get('Status')}
            """.strip(),
            'start': {'dateTime': start_datetime, 'timeZone': 'Asia/Kolkata'},
            'end': {'dateTime': end_datetime, 'timeZone': 'Asia/Kolkata'},
        }

        created_event = service.events().insert(calendarId='primary', body=event).execute()
        print(f"Event created: {created_event.get('htmlLink')}")

    except Exception as e:
        print(f"Error creating event: {e}")

def extract_valid_json(response_text):
    """
    Extract and validate JSON from the response text.
    """
    try:
        json_match = re.search(r"{.*}", response_text, re.DOTALL)
        return json.loads(json_match.group(0)) if json_match else {}
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        return {}

def extract_booking_details(email_content):
    """
    Extract booking details using Gemini AI with regex fallback.
    """
    try:
        # Gemini AI Extraction
        prompt = f"""
            Extract the following details as JSON:
            - Booking Reference, Title, Location, Travel Date
            - Lead Traveler Name, Hotel Pickup, Status

            Email content:
            {email_content}
        """
        response = model.generate_content(prompt)
        if response and hasattr(response, 'text') and response.text.strip():
            return extract_valid_json(response.text.strip())
    except Exception as e:
        print(f"Gemini AI extraction failed: {e}")

    # Regex Fallback
    print("Falling back to regex extraction.")
    regex_fields = {
        "Booking Reference": r"Booking Reference:\s*(#\S+)",
        "Title": r"Amended\s*(.*)",
        "Location": r"Location:\s*(.*)",
        "Travel Date": r"Travel Date:\s*(.*)",
        "Lead Traveler Name": r"Lead traveler name:\s*(.*)",
        "Hotel Pickup": r"Hotel Pickup:\s*(.*)",
    }
    return {field: re.search(pattern, email_content).group(1).strip() if re.search(pattern, email_content) else "Not Found" for field, pattern in regex_fields.items()}

def fetch_emails():
    """
    Fetch emails, extract booking details, and process them.
    """
    gmail = Gmail()
    query = construct_query(sender="Linda Tours Mumbai")
    try:
        messages = gmail.get_messages(query=query)
        if not messages:
            print("No messages found.")
            return

        print(f"Found {len(messages)} emails.")
        service = authenticate_calendar_api()

        for msg in messages[:10]:  # Process up to 10 emails
            email_content = msg.plain
            event_details = extract_booking_details(email_content)
            if event_details:
                results.append(event_details)
                create_event(service, event_details)
                send_data_to_website(event_details)
            else:
                print(f"Failed to extract details for email: {msg.id}")

    except Exception as e:
        print(f"Error fetching emails: {e}")

if __name__ == "__main__":
    fetch_emails()
from simplegmail import Gmail
from simplegmail.query import construct_query
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import re
import os
import pickle
import google.generativeai as genai
import json
import requests

# Google Gemini API Configuration
GENAI_API_KEY = "AIzaSyB6vX8Lddi4qMJlwnoiBxXEoII7gcwfw84"
genai.configure(api_key=GENAI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-pro-latest")

# API Endpoint
WEBSITE_API_URL = "http://localhost:3000/api/events"

results = []

def send_data_to_website(event_details):
    """
    Send event details to the API endpoint.
    """
    try:
        response = requests.post(WEBSITE_API_URL, json=event_details)
        if response.status_code == 200:
            print(f"Data sent successfully: {response.json()}")
        else:
            print(f"Failed to send data: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error while sending data: {e}")

def authenticate_calendar_api():
    """
    Authenticate and return Google Calendar API service instance.
    """
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('calendar', 'v3', credentials=creds)

def parse_travel_date(travel_date_str):
    """
    Parse the travel date from the given string using multiple formats.
    """
    formats = ['%a, %b %d, %Y', '%b %d, %Y', '%Y-%m-%d']
    for fmt in formats:
        try:
            return datetime.datetime.strptime(travel_date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date format: {travel_date_str}")

def delete_event(service, booking_reference, start_date, end_date):
    """
    Delete any existing events with the specified booking reference and dates.
    """
    try:
        service.events().delete(calendarId='primary', eventId=event_id).execute()
        events = service.events().list(
            calendarId='primary', q=booking_reference, singleEvents=True
        ).execute().get('items', [])

        for event in events:
            event_start = event['start'].get('dateTime') or event['start'].get('date')
            event_end = event['end'].get('dateTime') or event['end'].get('date')

            if start_date in event_start and end_date in event_end:
                service.events().delete(calendarId='primary', eventId=event['id']).execute()
                print(f"Deleted event: {event['summary']}")

    except Exception as e:
        print(f"Error deleting events: {e}")

def create_event(service, event_details):
    """
    Create a Google Calendar event from the extracted details.
    """
    try:
        # Parse travel date
        travel_date = parse_travel_date(event_details['Travel Date'])
        start_datetime = travel_date.replace(hour=9, minute=0).isoformat()
        end_datetime = travel_date.replace(hour=10, minute=0).isoformat()

        # Delete existing events before creating
        delete_event(service, event_details.get('Booking Reference', 'Unknown'), start_datetime, end_datetime)

        # Event body
        event = {
            'summary': event_details.get('Title', 'New Booking'),
            'description': f"""
                Booking Reference: {event_details.get('Booking Reference', 'Unknown')}
                Location: {event_details.get('Location')}
                Travel Date: {event_details['Travel Date']}
                Lead Traveler Name: {event_details.get('Lead Traveler Name')}
                Hotel Pickup: {event_details.get('Hotel Pickup')}
                Status: {event_details.get('Status')}
            """.strip(),
            'start': {'dateTime': start_datetime, 'timeZone': 'Asia/Kolkata'},
            'end': {'dateTime': end_datetime, 'timeZone': 'Asia/Kolkata'},
        }

        created_event = service.events().insert(calendarId='primary', body=event).execute()
        print(f"Event created: {created_event.get('htmlLink')}")

    except Exception as e:
        print(f"Error creating event: {e}")

def extract_valid_json(response_text):
    """
    Extract and validate JSON from the response text.
    """
    try:
        json_match = re.search(r"{.*}", response_text, re.DOTALL)
        return json.loads(json_match.group(0)) if json_match else {}
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        return {}

def extract_booking_details(email_content):
    """
    Extract booking details using Gemini AI with regex fallback.
    """
    try:
        # Gemini AI Extraction
        prompt = f"""
            Extract the following details as JSON:
            - Booking Reference, Title, Location, Travel Date
            - Lead Traveler Name, Hotel Pickup, Status

            Email content:
            {email_content}
        """
        response = model.generate_content(prompt)
        if response and hasattr(response, 'text') and response.text.strip():
            return extract_valid_json(response.text.strip())
    except Exception as e:
        print(f"Gemini AI extraction failed: {e}")

    # Regex Fallback
    print("Falling back to regex extraction.")
    regex_fields = {
        "Booking Reference": r"Booking Reference:\s*(#\S+)",
        "Title": r"Amended\s*(.*)",
        "Location": r"Location:\s*(.*)",
        "Travel Date": r"Travel Date:\s*(.*)",
        "Lead Traveler Name": r"Lead traveler name:\s*(.*)",
        "Hotel Pickup": r"Hotel Pickup:\s*(.*)",
    }
    return {field: re.search(pattern, email_content).group(1).strip() if re.search(pattern, email_content) else "Not Found" for field, pattern in regex_fields.items()}

def fetch_emails():
    """
    Fetch emails, extract booking details, and process them.
    """
    gmail = Gmail()
    query = construct_query(sender="Linda Tours Mumbai")
    try:
        messages = gmail.get_messages(query=query)
        if not messages:
            print("No messages found.")
            return

        print(f"Found {len(messages)} emails.")
        service = authenticate_calendar_api()

        for msg in messages[:10]:  # Process up to 10 emails
            email_content = msg.plain
            event_details = extract_booking_details(email_content)
            if event_details:
                results.append(event_details)
                create_event(service, event_details)
                send_data_to_website(event_details)
            else:
                print(f"Failed to extract details for email: {msg.id}")

    except Exception as e:
        print(f"Error fetching emails: {e}")

if __name__ == "__main__":
    fetch_emails()
