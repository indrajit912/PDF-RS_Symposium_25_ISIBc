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
    headers = ['Timestamp', 'Name', 'Email', 'Affiliation', 'Wanna talk?', 'Wanna attend?']
    print_sheet_data(sheet_data=data, headers=headers)


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
        affiliation = row[3].strip()
        wanna_talk = row[4].strip().lower() == "yes"

        if affiliation.lower() == "phd":
            total_phd += 1
            if wanna_talk:
                talk_phd += 1
        elif affiliation.lower() == "post doc":
            total_postdoc += 1
            if wanna_talk:
                talk_postdoc += 1

    total_responses = total_phd + total_postdoc
    total_speakers = talk_phd + talk_postdoc

    print(f"Total responses: ({total_phd} PhD, {total_postdoc} PostDoc) = {total_responses}")
    print(f"Wanna Talk: ({talk_phd} PhD, {talk_postdoc} PostDoc) = {total_speakers}")


def main():
    # Creating googleSheet
    client = GoogleSheetClient()
    spreadsheet = GoogleSheet(client=client, spreadsheet_id=SYMPOSIUM_SHEET_ID)
    ws = spreadsheet.get_worksheet(0)
    data = ws.get()

    print_responses_for_talk(data=data[1:])

    # Summarize responses
    summarize_responses(data=data[1:])
    


if __name__ == '__main__':
    main()
