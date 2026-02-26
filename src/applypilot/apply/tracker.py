import os
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1fapAJsuW66HLMLaTKUkpWCxFR86jNfSL5BJ3hxH7_L4'


def track_application(company: str, position: str, cover_letter: bool, location: str, in_office_amount: int,
                      posted_salary_top: int, website: str):
    """
    Track an application for a job with the given details to Google Sheets.

    Args:
        company (str): The name of the company.
        position (str): The position applied for.
        cover_letter (bool): Whether a cover letter was submitted.
        location (str): The location of the job in City, State Installs: ex. San Francisco, CA.
        in_office_amount (int): 0 for Remote, 1 for Hybrid, 2 for Full-Time.
        posted_salary_top (int): The top salary range posted for the job.
        website (str): The job posting website.
    Returns:
        str: The updated range in the Google Sheet if successful, None otherwise.
    """
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        logger.warning("GOOGLE_APPLICATION_CREDENTIALS not set, skipping Google Sheets tracking.")
        return None

    if not os.path.exists(creds_path):
        logger.error(f"Google credentials file not found at: {creds_path}")
        return None

    try:
        creds = service_account.Credentials.from_service_account_file(
            creds_path, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        # /Users/daktic/Downloads/_resume/generic/Garrett Berg - Software Engineer.pdf
        date = datetime.now().strftime("%m/%d/%Y")
        row = []
        row.append(date)
        row.append(company)
        row.append(position)
        row.append("Yes" if cover_letter else "No")
        row.append("No")
        if not location:
            row.append("Remote")
        else:
            row.append(location)

        if in_office_amount == 1:
            row.append("Hybrid")
        elif in_office_amount == 2:
            row.append("Full-Time")
        else:
            row.append("N/A")

        row.append(posted_salary_top)
        row.append("No")
        row.append("")
        row.append("Applied with ApplyPilot")
        row.append(website)
        body = {
            'values': [row]
        }
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range="Jobs Applied For!A1",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        updated_range = result.get('updates', {}).get('updatedRange')
        logger.info(f"Successfully tracked application to Google Sheets: {updated_range}")
        return updated_range

    except HttpError as err:
        logger.error(f"Google Sheets API error: {err}")
    except Exception as e:
        logger.error(f"Unexpected error during Google Sheets tracking: {e}")
    return None

