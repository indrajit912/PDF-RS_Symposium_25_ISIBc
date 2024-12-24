# schedule.py
# Author: Indrajit Ghosh
# Created On: Dec 24, 2024
#
import random

from Google import *
from main import get_participants_and_speakers, SYMPOSIUM_SHEET_ID


def create_schedule(speakers, total_days, max_talks_per_day, randomize=False):
    """
    Creates a schedule for the speakers over the given number of days.

    This function distributes speakers across the specified number of days, ensuring that each day 
    has an equal or nearly equal number of talks, while respecting the maximum number of talks per day.
    Optionally, the schedule can be randomized.

    :param speakers: List of speakers (list of lists) containing speaker details. 
                      Each inner list represents a speaker with their information (e.g., name, talk title).
    :param total_days: The total number of days to schedule talks. 
                        The schedule will distribute talks evenly across these days.
    :param max_talks_per_day: The maximum number of talks allowed per day. 
                               If there are more talks than this limit, they will be distributed across days.
    :param randomize: If True, the allocation of speakers to days will be randomized. Defaults to False.
    :return: A list of lists representing the schedule, where each list corresponds to a day and contains 
             the speakers scheduled on that day. The format is [[day_1_talks], [day_2_talks], ..., [day_n_talks]].
    """
    total_speakers = len(speakers)
    schedule = [[] for _ in range(total_days)]

    # Randomize the list of speakers if requested
    if randomize:
        random.shuffle(speakers)

    # Determine how to distribute the talks
    speakers_per_day = total_speakers // total_days
    remaining_speakers = total_speakers % total_days

    current_speaker = 0

    for day in range(total_days):
        # Distribute speakers per day, adding extra ones to the last days if necessary
        talks_today = speakers_per_day + (1 if day < remaining_speakers else 0)
        
        # Add speakers for today (up to max_talks_per_day)
        for _ in range(min(talks_today, max_talks_per_day)):
            if current_speaker < total_speakers:
                schedule[day].append(speakers[current_speaker])
                current_speaker += 1

    return schedule


def main():
    """
    Main function that orchestrates the scheduling of talks.

    This function fetches data from a Google Sheets document, processes the participant and speaker 
    information, creates a schedule for the talks, and prints the schedule in a readable format.

    The function does the following:
    1. Fetches data from a Google Sheets document.
    2. Extracts the speaker details from the data.
    3. Generates a schedule with a maximum of 5 talks per day over 3 days.
    4. Prints the schedule in a human-readable format.

    :return: None
    """
    # Fetching data from Google Sheets
    client = GoogleSheetClient()
    spreadsheet = GoogleSheet(client=client, spreadsheet_id=SYMPOSIUM_SHEET_ID)
    ws = spreadsheet.get_worksheet(0)
    data = ws.get()

    if not data or len(data) < 2:
        print("No data available in the sheet.")
        return
    
    _, speakers = get_participants_and_speakers(data)
    # Create a schedule
    schedule = create_schedule(speakers, total_days=3, max_talks_per_day=5, randomize=False)

    # Print the schedule
    for day, talks in enumerate(schedule, start=1):
        print(f"Day {day}:")
        for talk in talks:
            print(f"  - {talk[1]} ({talk[3]})")
        print("\n")

if __name__ == "__main__":
    main()
