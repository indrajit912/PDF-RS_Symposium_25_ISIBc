# Tests
#
# Author: Indrajit Ghosh
#
# Date: Aug 28, 2022
#

from pathlib import Path
from model import *


def main():
    # email = GmailMessage(
    #     sender_email_id='ma19d002@smail.iitm.ac.in',
    #     to=['indrajitghosh912@gmail.com', 'rs_math1902@isibang.ac.in'],
    #     subject='EmailMessage class Testing',
    #     cc='indrajitghosh912@outlook.com',
    #     bcc='indrajitghosh2014@gmail.com',
    #     email_plain_text="Hi,\nThis is a system generated email. Kindly don't reply to this!",
    #     attachments=['./spam/mmath.xlsx']
    # )

    # print(email)
    # email.send()

    INDRAJIT_DELL_ID = "1UZoQ6La1Pw2H4apxrU9RGLlw9PM4ySJO"

    drive = GoogleDrive()

    files = drive.get_filelist(INDRAJIT_DELL_ID)
    pprint(files)

    # sh_id = '1oAhuWvGrdAnSnPqOuDWPFTyqJzS4ihVVOqWfqLDC86E'

    # gsheet_client = GoogleSheetClient()
    # spreadsheet = GoogleSheet(client=gsheet_client, spreadsheet_id=sh_id)

    # ws = spreadsheet.get_worksheet_by_title("India")
    # print(ws)

    BOT_CRED_JSON = Path.home() / "Downloads" / "credentials.json"
    BOT_TOKEN_JSON = Path.home() / "Downloads" / "token.json"

    # download_bookcover_imgs(where_to_save_imgs=BOOK_COVER_IMG_DIR)
    # pprint(get_list_of_bookcoverimg_ids_from_GDrive())

    # drive = GoogleDrive(
    #     authorized_user_file=BOT_CRED_JSON,
    #     client_secret_file=BOT_TOKEN_JSON
    # )






if __name__ == '__main__':
    main()