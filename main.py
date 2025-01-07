# main.py
# Author: Indrajit Ghosh
# Created On: Dec 20, 2024
#
# This script can be used to generate all necessary 
# documents for the Symposium
#
from Google import *
from tabulate import tabulate

SYMPOSIUM_SHEET_ID = "1oWc3pIBOhDinA3qak3Hqt8YxkubCc17sIIsZuqeUEPM"
SPEAKER_INFO_SHEET_ID = "1Nh6gFkamgS06SFGWfhMWhUmc1CgPZLVFYKbtPBSsrck"


def print_sheet_data(sheet_data, headers=None):
    """
    Prints Google Sheet data in a tabular format.

    :param sheet_data: List of lists representing the rows of the sheet.
    :param headers: Optional list of headers for the table.
    """
    if headers:
        print(tabulate(sheet_data, headers=headers, tablefmt="grid"))
    else:
        print(tabulate(sheet_data, tablefmt="grid"))


def print_responses_for_talk(data):
    """
    Prints responses for the talk in a tabular format.

    :param data: List of lists representing the rows of the sheet.
    """
    headers = ['Timestamp', 'Name', 'Email', 'Affiliation', 'Wanna talk?', 'Wanna attend?']
    print_sheet_data(sheet_data=data, headers=headers)


def get_participants_and_speakers(data):
    """
    Extracts participants and speakers from the Google Sheet data.

    :param data: List of lists representing the rows of the sheet (excluding headers).
    :return: Tuple containing two lists:
             - participants: List of [timestamp, name, email, affiliation] for attendees.
             - speakers: List of [timestamp, name, email, affiliation] for those who want to give a talk.
    """
    participants = []
    speakers = []

    for row in data:
        timestamp = row[0].strip()
        name = row[1].strip()
        email = row[2].strip()
        affiliation = row[3].strip()
        wanna_talk = row[4].strip().lower() == "yes"
        wanna_attend = row[5].strip().lower() == "yes"

        if wanna_attend:
            participants.append([timestamp, name, email, affiliation])
        if wanna_talk:
            speakers.append([timestamp, name, email, affiliation])

    return participants, speakers


def summarize_responses(data):
    """
    Summarizes the total responses and the "Wanna Talk" preferences by category (PhD, PostDoc).

    :param data: List of lists representing the rows of the sheet (excluding headers).
    """
    total_phd = 0
    total_postdoc = 0
    talk_phd = 0
    talk_postdoc = 0

    for row in data:
        affiliation = row[3].strip().lower()
        wanna_talk = row[4].strip().lower() == "yes"

        if affiliation == "phd":
            total_phd += 1
            if wanna_talk:
                talk_phd += 1
        elif affiliation == "post doc":
            total_postdoc += 1
            if wanna_talk:
                talk_postdoc += 1

    total_responses = total_phd + total_postdoc
    total_speakers = talk_phd + talk_postdoc

    print(f"Total responses: ({total_phd} PhD, {total_postdoc} PostDoc) = {total_responses}")
    print(f"Wanna Talk: ({talk_phd} PhD, {talk_postdoc} PostDoc) = {total_speakers}")


def all_data():
    # Fetching data from Google Sheets
    client = GoogleSheetClient()
    spreadsheet = GoogleSheet(client=client, spreadsheet_id=SYMPOSIUM_SHEET_ID)
    ws = spreadsheet.get_worksheet(0)
    data = ws.get()

    if not data or len(data) < 2:
        print("No data available in the sheet.")
        return

    # Summarize responses
    summarize_responses(data=data[1:])

    # Print responses for the talk
    print_responses_for_talk(data=data[1:])

    # Get participants and speakers
    participants, speakers = get_participants_and_speakers(data=data[1:])
    print(f"\nParticipants: {len(participants)}")
    print_sheet_data(sheet_data=participants, headers=["Timestamp", "Name", "Email", "Affiliation"])
    print(f"\nSpeakers: {len(speakers)}")
    print_sheet_data(sheet_data=speakers, headers=["Timestamp", "Name", "Email", "Affiliation"])

from tabulate import tabulate

def speaker_info():
    """
    Fetches and prints the speaker information from the Google Sheet in a two-column format using tabulate.
    """
    # Initialize Google Sheet Client
    speaker_client = GoogleSheetClient()
    sh = GoogleSheet(client=speaker_client, spreadsheet_id=SPEAKER_INFO_SHEET_ID)
    worksheet = sh.get_worksheet(0)
    speaker_data = worksheet.get()

    if not speaker_data or len(speaker_data) < 2:
        print("No speaker information available in the sheet.")
        return

    # Extract headers and data
    headers = speaker_data[0]
    data = speaker_data[1:]

    print("\n================= Speaker Information =================\n")
    
    # Format data into two columns: header and corresponding value
    formatted_data = []
    for row in data:
        speaker_info = list(zip(headers, row))
        formatted_data.append(speaker_info)

    # Use tabulate to display data in a neat table format (2 columns)
    for idx, speaker in enumerate(formatted_data, start=1):
        print(f"\nSpeaker {idx}:")
        print(tabulate(speaker, tablefmt="fancy_grid", stralign="left"))


def main():
    speaker_info()


if __name__ == '__main__':
    main()
