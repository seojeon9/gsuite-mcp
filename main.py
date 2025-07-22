#!/usr/bin/env python3
"""
Google Workspace MCP Server using FastMCP
"""

import os
import json
import base64
from datetime import datetime
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables from .env file
load_dotenv()

# Environment variables required for OAuth
CLIENT_ID = os.getenv("GOOGLE_APP_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_APP_CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("GOOGLE_APP_REFRESH_TOKEN")

if not all([CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN]):
    raise ValueError("Required Google OAuth credentials not found in environment variables")

# Set up OAuth2 credentials
credentials = Credentials(
    token=None,
    refresh_token=REFRESH_TOKEN,
    token_uri="https://oauth2.googleapis.com/token",
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET,
    scopes=[
        "https://www.googleapis.com/auth/gmail.modify",
        "https://www.googleapis.com/auth/calendar.events"
    ]
)

# Initialize API clients
gmail_service = build("gmail", "v1", credentials=credentials)
calendar_service = build("calendar", "v3", credentials=credentials)

# Initialize FastMCP
mcp = FastMCP("google-workspace-server")

def get_email_body(payload: Dict[str, Any]) -> str:
    """Extract email body from Gmail message payload"""
    if not payload:
        return ""

    if payload.get("body") and payload["body"].get("data"):
        return base64.urlsafe_b64decode(payload["body"]["data"]).decode("utf-8")

    if payload.get("parts"):
        for part in payload["parts"]:
            if part.get("mimeType") == "text/plain" and part.get("body", {}).get("data"):
                return base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8")

    return "(No body content)"

@mcp.tool()
def list_emails(maxResults: int = 10, query: str = "") -> str:
    """List recent emails from Gmail inbox"""
    try:
        response = gmail_service.users().messages().list(
            userId="me",
            maxResults=maxResults,
            q=query
        ).execute()

        messages = response.get("messages", [])
        email_details = []

        for msg in messages:
            detail = gmail_service.users().messages().get(
                userId="me",
                id=msg["id"]
            ).execute()

            headers = detail["payload"].get("headers", [])
            subject = next((h["value"] for h in headers if h["name"] == "Subject"), "")
            from_addr = next((h["value"] for h in headers if h["name"] == "From"), "")
            date = next((h["value"] for h in headers if h["name"] == "Date"), "")
            body = get_email_body(detail["payload"])

            email_details.append({
                "id": msg["id"],
                "subject": subject,
                "from": from_addr,
                "date": date,
                "body": body,
            })

        return json.dumps(email_details, indent=2)

    except Exception as e:
        return f"Error fetching emails: {str(e)}"

@mcp.tool()
def search_emails(query: str, maxResults: int = 10) -> str:
    """Search emails with advanced query"""
    try:
        response = gmail_service.users().messages().list(
            userId="me",
            maxResults=maxResults,
            q=query
        ).execute()

        messages = response.get("messages", [])
        email_details = []

        for msg in messages:
            detail = gmail_service.users().messages().get(
                userId="me",
                id=msg["id"]
            ).execute()

            headers = detail["payload"].get("headers", [])
            subject = next((h["value"] for h in headers if h["name"] == "Subject"), "")
            from_addr = next((h["value"] for h in headers if h["name"] == "From"), "")
            date = next((h["value"] for h in headers if h["name"] == "Date"), "")
            body = get_email_body(detail["payload"])

            email_details.append({
                "id": msg["id"],
                "subject": subject,
                "from": from_addr,
                "date": date,
                "body": body,
                "labels": detail.get("labelIds", []),
            })

        return json.dumps(email_details, indent=2)

    except Exception as e:
        return f"Error fetching emails: {str(e)}"

@mcp.tool()
def send_email(to: str, subject: str, body: str, cc: str = None, bcc: str = None) -> str:
    """Send a new email"""
    try:
        headers = [
            'Content-Type: text/plain; charset="UTF-8"',
            "MIME-Version: 1.0",
            f"To: {to}",
        ]

        if cc:
            headers.append(f"Cc: {cc}")
        if bcc:
            headers.append(f"Bcc: {bcc}")

        headers.append(f"Subject: {subject}")

        email = "\r\n".join(headers) + "\r\n\r\n" + body

        # Encode in base64url
        encoded_message = base64.urlsafe_b64encode(email.encode()).decode().rstrip("=")

        response = gmail_service.users().messages().send(
            userId="me",
            body={"raw": encoded_message}
        ).execute()

        return f"Email sent successfully. Message ID: {response['id']}"

    except Exception as e:
        return f"Error sending email: {str(e)}"

@mcp.tool()
def modify_email(id: str, addLabels: List[str] = None, removeLabels: List[str] = None) -> str:
    """Modify email labels (archive, trash, mark read/unread)"""
    try:
        addLabels = addLabels or []
        removeLabels = removeLabels or []

        response = gmail_service.users().messages().modify(
            userId="me",
            id=id,
            body={
                "addLabelIds": addLabels,
                "removeLabelIds": removeLabels,
            }
        ).execute()

        return f"Email modified successfully. Updated labels for message ID: {response['id']}"

    except Exception as e:
        return f"Error modifying email: {str(e)}"

@mcp.tool()
def list_events(maxResults: int = 10, timeMin: str = None, timeMax: str = None) -> str:
    """List calendar events. You can specify timeMin and timeMax parameters to fetch events from past dates or within a specific date range. Both parameters accept ISO format datetime strings."""
    try:
        if not timeMin:
            timeMin = datetime.now().isoformat() + "Z"

        request_params = {
            "calendarId": "primary",
            "timeMin": timeMin,
            "maxResults": maxResults,
            "singleEvents": True,
            "orderBy": "startTime",
        }

        if timeMax:
            request_params["timeMax"] = timeMax

        response = calendar_service.events().list(**request_params).execute()

        events = []
        for event in response.get("items", []):
            events.append({
                "id": event["id"],
                "summary": event.get("summary"),
                "start": event.get("start"),
                "end": event.get("end"),
                "location": event.get("location"),
            })

        return json.dumps(events, indent=2)

    except Exception as e:
        return f"Error fetching calendar events: {str(e)}"

@mcp.tool()
def create_event(summary: str, start: str, end: str, location: str = None,
                description: str = None, attendees: List[str] = None) -> str:
    """Create a new calendar event"""
    try:
        event = {
            "summary": summary,
            "start": {
                "dateTime": start,
                "timeZone": "Asia/Seoul",
            },
            "end": {
                "dateTime": end,
                "timeZone": "Asia/Seoul",
            },
        }

        if location:
            event["location"] = location
        if description:
            event["description"] = description
        if attendees:
            event["attendees"] = [{"email": email} for email in attendees]

        response = calendar_service.events().insert(
            calendarId="primary",
            body=event
        ).execute()

        return f"Event created successfully. Event ID: {response['id']}"

    except Exception as e:
        return f"Error creating event: {str(e)}"

@mcp.tool()
def update_event(eventId: str, summary: str = None, location: str = None,
                description: str = None, start: str = None, end: str = None,
                attendees: List[str] = None) -> str:
    """Update an existing calendar event"""
    try:
        event = {}

        if summary:
            event["summary"] = summary
        if location:
            event["location"] = location
        if description:
            event["description"] = description
        if start:
            event["start"] = {
                "dateTime": start,
                "timeZone": "Asia/Seoul",
            }
        if end:
            event["end"] = {
                "dateTime": end,
                "timeZone": "Asia/Seoul",
            }
        if attendees:
            event["attendees"] = [{"email": email} for email in attendees]

        response = calendar_service.events().patch(
            calendarId="primary",
            eventId=eventId,
            body=event
        ).execute()

        return f"Event updated successfully. Event ID: {response['id']}"

    except Exception as e:
        return f"Error updating event: {str(e)}"

@mcp.tool()
def delete_event(eventId: str) -> str:
    """Delete a calendar event"""
    try:
        calendar_service.events().delete(
            calendarId="primary",
            eventId=eventId
        ).execute()

        return f"Event deleted successfully. Event ID: {eventId}"

    except Exception as e:
        return f"Error deleting event: {str(e)}"

if __name__ == "__main__":
    mcp.run(transport="stdio")