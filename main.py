from __future__ import print_function
import os.path
import pytz
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from pprint import pprint
from datetime import datetime
from time import sleep

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/youtube']


def auth():
    """Receive the creds for the google service."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


class Youtube:
    def __init__(self):
        self.service = build('youtube', 'v3', credentials=auth())

    def get_video_views(self, video_id, t_language):
        results = self.service.videos().list(
            part="statistics",
            id=video_id,
            hl=t_language
        ).execute()
        return int(results["items"][0]["statistics"]["viewCount"])


class Sheets:
    def __init__(self, t_spreadsheet_id):
        self.service = build('sheets', 'v4', credentials=auth())
        self.spreadsheet_id = t_spreadsheet_id

        first_mv_col = "C"
        second_mv_col = "D"
        self.rate_per_hr = '=(INDIRECT(CONCAT("E", ROW())) - INDIRECT(CONCAT("E", ROW() - 1)))/' \
                           '(INDIRECT(CONCAT("B", ROW())) - INDIRECT(CONCAT("B", ROW() - 1)))'
        self.summed_views = f"={first_mv_col}:{first_mv_col}+{second_mv_col}:{second_mv_col}"
        self.cumulative_rate = '=(INDIRECT(CONCAT("E", ROW())) / INDIRECT(CONCAT("B", ROW())))'
        self.sheet_range = "A:A"
        self.current_hour = '=(INDIRECT(CONCAT("B", ROW() - 1)) + 1)'
        self.kst_timezone = pytz.timezone("Asia/Seoul")

    def add_views(self, views: int, views_2: int):

        values = [[datetime.now(self.kst_timezone).strftime('%Y-%m-%d %H:%M:%S'),
                   self.current_hour,
                   f"{views:,}",
                   f"{views_2:,}",
                   self.summed_views,
                   self.rate_per_hr,
                   self.cumulative_rate]]

        request = self.service.spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id,
            range=self.sheet_range,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": values}
        )
        response = request.execute()
        pprint(response)


if __name__ == '__main__':
    # constants
    spreadsheet_id = "1dDgCATWYKDinbgTLLI26PBmDOfSuONnpQW7MSQ_TNr0"
    first_video_id = "I67jpVeNCro"
    second_video_id = "EtiPbWzUY9o"
    language = "en-us"
    sleep_time = 30*60

    youtube = Youtube()
    sheets = Sheets(spreadsheet_id)

    """
    Excel Table Format is Expected to be:
    
    KST TimeStamp, Hours,	Official,	Genie, 	Total Views,	rate of that hour,	cumulative rate
    """
    while True:
        try:
            first_video_views = 0 if not first_video_id else youtube.get_video_views(first_video_id, language)
            second_video_views = 0 if not second_video_id else youtube.get_video_views(second_video_id, language)
            sheets.add_views(first_video_views, second_video_views)
            print(f"Updated Video 1 with {first_video_views} views.")
            print(f"Updated Video 2 with {second_video_views} views.")
            print("Sleeping for 30 minutes.")
            sleep(sleep_time)
        except Exception as e:
            print(e)
