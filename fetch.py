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
import time

GENAI_API_KEY = "AIzaSyByZukGhz0fp6gftgrrmwbUB4BymROAp84"
genai.configure(api_key=GENAI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-pro-latest")

results = []

WEBSITE_API_URL = "http://localhost:3000/api/events"  # Replace with your API endpoint

def main():
    """
    Main function to fetch emails, process them, and create/update events.
    Runs once per hour.
    """
    while True:  # Infinite loop to run the code repeatedly
        try:
            print("Starting email fetch and event processing...")
            fetch_emails()  # Call your existing function
            print("Processing completed. Waiting for the next run...")
        except Exception as e:
            print(f"An error occurred during processing: {e}")
        
        # Wait for one hour (3600 seconds) before running again
        time.sleep(3600)

def send_data_to_website(event_details):
    """
    Send extracted event details to the website API.
    """
    try:
        response = requests.post(WEBSITE_API_URL, json=event_details)
        if response.status_code == 200:
            print(f"Data successfully sent to website: ")
        else:
            print(f"Failed to send data to website. Status code: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error sending data to website: {e}")

def authenticate_calendar_api():
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
    service = build('calendar', 'v3', credentials=creds)
    return service

def parse_travel_date(travel_date_str):
    """
    Attempt to parse the travel date from the provided string using multiple formats.
    """
    formats = [
        '%a, %b %d, %Y',  # Example: Tue, Feb 18, 2025
        '%b %d, %Y',       # Example: Feb 18, 2025
        '%Y-%m-%d'         # Example: 2025-02-18 (ISO format)
    ]
    for fmt in formats:
        try:
            return datetime.datetime.strptime(travel_date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Date format not recognized: {travel_date_str}")

def update_event_to_canceled(service, booking_reference):
    """
    Update the event's title and status to "Canceled" if the event with the given Booking Reference exists.
    """
    try:
        print("inside update")
        # Search for an existing event with the same Booking Reference
        existing_events = service.events().list(
            calendarId='primary',
            singleEvents=True,
            q=booking_reference,
        ).execute().get('items', [])

        if existing_events:
            for event in existing_events:
                if "confirmed" in event.get('description', ''):
                    event['summary'] = f"Canceled - {event['summary']}"
                    event['description'] += f"\nUpdated Status: Canceled"
                    service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()
                    print(f"Updated event to 'Canceled': {event.get('htmlLink')}")
        else:
            print(f"No event found for Booking Reference '{booking_reference}'.")
    except Exception as e:
        print(f"Error updating event to 'Canceled': {e}")

def create_event(service, event_details_list):
    """
    Process multiple event details and create or update events dynamically.
    Handles "Confirmed" and "Canceled" statuses.
    """
    try:
        if isinstance(event_details_list, dict):
            event_details_list = [event_details_list]
        elif isinstance(event_details_list, str):
            event_details_list = json.loads(event_details_list)

        if not isinstance(event_details_list, list):
            raise ValueError("event_details_list must be a list of dictionaries.")

        for event_details in event_details_list:
            if not isinstance(event_details, dict):
                print(f"Skipping invalid event details: {event_details}")
                continue

            print(f"Processing event: {event_details}")

            booking_reference = event_details.get('Booking Reference', 'Not Provided')

            # Check the event status
            status = event_details.get('Status', '').lower()
            print(f"Checking event status for {booking_reference}: {status}")  # <-- Add this line

            if status == 'cancelled':
                print("helo in if")
                # If event is canceled, check and update the event
                update_event_to_canceled(service,booking_reference)
                continue

            if status == 'amended':
                print(f"Skipping event creation for status '{status}' with Booking Reference '{booking_reference}'.")
                continue

            # Check if an event with the same Booking Reference already exists
            existing_events = service.events().list(
                calendarId='primary',
                singleEvents=True,
                q=booking_reference,
            ).execute().get('items', [])

            if existing_events:
                print(f"Event with Booking Reference '{booking_reference}' already exists. Skipping creation.")
                continue

            # Parse the Travel Date and calculate the event start and end datetime
            travel_date_str = event_details.get('Travel Date', None)
            if not travel_date_str:
                print(f"Missing Travel Date for Booking Reference '{booking_reference}'. Skipping event.")
                continue

            try:
                travel_date = parse_travel_date(travel_date_str)
                start_datetime = travel_date.replace(hour=9, minute=0).isoformat()
                end_datetime = travel_date.replace(hour=10, minute=0).isoformat()
            except ValueError as e:
                print(f"Error parsing Travel Date '{travel_date_str}' for Booking Reference '{booking_reference}': {e}")
                continue

            # Build the event description
            description = f"""
            Booking Reference: {event_details.get('Booking Reference', 'Not Provided')}
            Title: {event_details.get('Title', 'New Booking')}
            Location: {event_details.get('Location', 'Not Provided')}
            Travel Date: {event_details.get('Travel Date', 'Not Provided')}
            Lead Traveler Name: {event_details.get('Lead Traveler Name', 'Not Provided')}
            Hotel Pickup: {event_details.get('Hotel Pickup', 'Not Provided')}
            Status: {event_details.get('Status', 'Not Provided')}
            """

            # Create the event body
            event = {
                'summary': event_details.get('Title', 'New Booking'),
                'description': description.strip(),
                'start': {
                    'dateTime': start_datetime,
                    'timeZone': 'Asia/Kolkata',
                },
                'end': {
                    'dateTime': end_datetime,
                    'timeZone': 'Asia/Kolkata',
                },
            }

            created_event = service.events().insert(calendarId='primary', body=event).execute()
            print(f"Event created: {created_event.get('htmlLink')}")

    except Exception as e:
        print(f"Error creating events: {e}")


def extract_valid_json(response_text):
    try:
        json_match = re.search(r"{.*}", response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(0))
        else:
            print("Error: No valid JSON found in the response.")
            return {}
    except json.JSONDecodeError as e:
        print(f"Error parsing the JSON response: {e}")
        return {}

def extract_booking_details_with_genai(email_content):
    prompt = f"""Extract the following event details from the given text:
    - Booking Reference
    - Title
    - Location
    - Travel Date
    - Lead Traveler Name
    - Hotel Pickup
    - Status
    - Start DateTime
    - End DateTime

    Return the event details as a valid JSON object in this format:
    {{
      "Booking Reference": "<reference>",
      "Title": "<title>",
      "Location": "<location>",
      "Travel Date": "<travel_date>",
      "Lead Traveler Name": "<lead_traveler>",
      "Hotel Pickup": "<hotel_pickup>",
      "Status": "<status>",
      "Start DateTime": "<start_datetime>",
      "End DateTime": "<end_datetime>"
    }}
    Ensure the response is strictly in JSON format without extra commentary.
    
    Email content:
    {email_content}
    """

    try:
        response = model.generate_content(prompt)

        if response and hasattr(response, 'text') and response.text.strip():
            raw_response = response.text.strip()
            return extract_valid_json(raw_response)  
        else:
            print("Error: Empty or invalid response from Gemini AI.")
            return None

    except Exception as e:
        print(f"Error extracting details with Gemini AI: {e}")
        return None

def extract_booking_details_with_regex(email_content):
    try:
        booking_ref = re.search(r"Booking Reference:\s*(#\S+)", email_content)
        tour_name = re.search(r"Amended\s*(.*)", email_content)
        location = re.search(r"Location:\s*(.*)", email_content)
        travel_date = re.search(r"Travel Date:\s*(.*)", email_content)
        lead_traveler = re.search(r"Lead traveler name:\s*(.*)", email_content)
        hotel_pickup = re.search(r"Hotel Pickup:\s*(.*)", email_content)
        
        details = {
            "Booking Reference": booking_ref.group(1) if booking_ref else "Not Found",
            "Title": tour_name.group(1).strip() if tour_name else "Not Found",
            "Location": location.group(1).strip() if location else "Not Found",
            "Travel Date": travel_date.group(1).strip() if travel_date else "Not Found",
            "Lead Traveler Name": lead_traveler.group(1).strip() if lead_traveler else "Not Found",
            "Hotel Pickup": hotel_pickup.group(1).strip() if hotel_pickup else "Not Found",
        }
        
        try:
            travel_datetime = datetime.datetime.strptime(details["Travel Date"], "%a, %b %d, %Y")
            details["start_datetime"] = travel_datetime.isoformat()
            details["end_datetime"] = (travel_datetime + datetime.timedelta(hours=4)).isoformat()
        except Exception as e:
            print(f"Error parsing travel date: {e}")
            details["start_datetime"] = None
            details["end_datetime"] = None
        
        return details

    except Exception as e:
        print(f"Error extracting booking details: {e}")
        return None

def fetch_emails():
    gmail = Gmail()  
    query = construct_query(sender="Linda Tours Mumbai")
    try:
        messages = gmail.get_messages(query=query)
        if not messages:
            print("No messages found.")
            return
        print(f"Found {len(messages)} emails.")

        for msg in messages[:10]:  # Process up to 10 emails
            email_content = msg.plain

            # Extract details using Gemini AI
            extracted_details = extract_booking_details_with_genai(email_content)
            if extracted_details:
                results.append(extracted_details)
            else:
                print("Falling back to regex parsing.")
                # Extract details using regex
                extracted_details = extract_booking_details_with_regex(email_content)
                if extracted_details:
                    results.append(extracted_details)
                else:
                    print(f"Failed to extract details from email: {msg.subject}")

        # Send extracted data to the website API
        for event in results:
            send_data_to_website(event)

        # Authenticate and create or update events
        service = authenticate_calendar_api()
        create_event(service, results)

    except Exception as e:
        print(f"Error fetching emails: {e}")


if __name__ == "__main__":
    fetch_emails()
    main()
