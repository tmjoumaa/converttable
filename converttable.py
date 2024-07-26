from dataclasses import dataclass
from typing import List
import sys

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SHEET_ID = 'xxx'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

@dataclass
class Person:
    name: str
    organization: str
    sessions: str

people: List[Person] = []

def findfirstletter(s: str) -> int: 
    for i, c in enumerate(s):
        if c.isalpha():
            return i
    return -1

def findfirstwhitespace(s: str) -> int:
    for i, c in enumerate(s):
        if c.isspace():
            return i
    return -1

def main():
    infile = None
    startnamedecode = False
    people = []
    sheeturl = None

    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        sheeturl = sys.argv[1]
        infile = sys.argv[2]
    except IndexError:
        print("Usage: python3 converttable.py <Google Sheet URL> <filename>")
        quit(0x41)

    SHEET_ID = sheeturl.split("/")[5]

    with open(infile, 'r') as f:
        lines = f.readlines()
        for line in lines:
            if "------" in line:
                    startnamedecode = True
                    continue
            if startnamedecode:
                # Find first large space
                fls = line.find("   ")

                if fls == -1:
                        print("Error: line not formatted correctly")
                        break

                name = line[:fls]

                remaining = line[fls:]
                remaining = remaining[findfirstletter(remaining):]

                org = remaining[:remaining.find("   ")]

                s = remaining = remaining[remaining.find("   "):].replace(" ", "").replace("\n", "")
                p = Person(name, org, s)
                people.append(p)
    print(f"Constructed {len(people)} people")

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        response_date = sheet.values().update(
        spreadsheetId=SHEET_ID,
        valueInputOption='RAW',
        range=f"A1:C{len(people)}",
        body=dict(
            majorDimension='COLUMNS',
            values=[[i.name for i in people], [i.organization for i in people], [i.sessions for i in people]])
    ).execute()

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()