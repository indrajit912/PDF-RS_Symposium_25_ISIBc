# Classes required for Google APIs
#
# Author: Indrajit Ghosh
#
# Date: Aug 28, 2022
#


import os, io, base64
import os.path
from pathlib import Path
from pprint import pprint
from bs4 import BeautifulSoup

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload

from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate, COMMASPACE
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email import encoders
import mimetypes

from google.oauth2 import service_account

from .exceptions import *


class EmailMessage(MIMEMultipart):
    """
    A class representing an email message

    Author: Indrajit Ghosh

    Date: Aug 28, 2022

    Parameters:
    -----------
        `sender_email_id`: `str`
        `to`: `str`
        `subject`: `str`
        `email_plain_text`: `str`; (This is the email body)
        `email_html_text` : `str`; (If you want to add some html text use this)
        `cc`: `str` / [`str`, `str`, ..., `str`]; (Carbon Copy)
        `bcc`: `str` / [`str`, `str`, ..., `str`]; (Blind Carbon Copy)
        `attachments` : `str` / [`str`, `str`, ..., `str`]; (Attachments)

    Returns:
    --------
        `MIMEMultipart` object (An Email Message)
    """

    def __init__(
        self,
        sender_email_id:str,
        to,
        subject:str=None,
        email_plain_text:str=None,
        email_html_text:str=None,
        cc=None,
        bcc=None,
        attachments=None
    ):

        cc = [] if cc is None else cc
        bcc = [] if bcc is None else bcc
        attachments = [] if attachments is None else attachments
        attachments = [attachments] if isinstance(attachments, str) else attachments

        if not isinstance(cc, list):
            cc = [cc]
        if not isinstance(bcc, list):
            bcc = [bcc]

        # Attributes
        self.sender = sender_email_id
        self.to = [to] if isinstance(to, str) else to
        self.cc = cc
        self.bcc = bcc
        self.recipients = self.to + self.cc + self.bcc
        self.subject = subject
        self.plain_msg = email_plain_text
        self.html_msg = email_html_text
        self.attachments = attachments

        MIMEMultipart.__init__(self)

        # Structure email
        if sender_email_id == "ma19d002@smail.iitm.ac.in":
            self['From'] = formataddr(("Indrajit Ghosh (SRF, SMU, ISIBc)", sender_email_id))
        elif sender_email_id == "indrajitsbot@gmail.com":
            self['From'] = formataddr(("Indrajit Ghosh (SRF, SMU, ISIBc)", sender_email_id))
        else:
            self['From'] = self.sender

        self['To'] = COMMASPACE.join(self.to)
        self['Cc'] = COMMASPACE.join(self.cc) if cc != [] else ''
        self['Bcc'] = COMMASPACE.join(self.bcc) if bcc != [] else ''
        self['Date'] = formatdate(localtime=True)
        self['Subject'] = subject

        # Attaching email text
        if self.plain_msg is not None:
            self.attach(MIMEText(self.plain_msg, 'plain'))
        if self.html_msg is not None:
            self.attach(MIMEText(self.html_msg, 'html'))

        # Adding attachments
        self.add_attachments()
        
    
    def add_attachments(self):
        
        for attached_file in self.attachments:

            attached_file = Path(attached_file)

            my_mimetype, encoding = mimetypes.guess_type(attached_file)

            if my_mimetype is None or encoding is not None:
                my_mimetype = 'application/octet-stream' 


            main_type, sub_type = my_mimetype.split('/', 1)# split only at the first '/'

            #this part is used to tell how the file should be read and stored (r, or rb, etc.)
            if main_type == 'text':
                print("text attached")
                temp = open(attached_file, 'r')  # 'rb' will send this error: 'bytes' object has no attribute 'encode'
                attachment = MIMEText(temp.read(), _subtype=sub_type)
                temp.close()

            elif main_type == 'image':
                print("image attached")
                temp = open(attached_file, 'rb')
                attachment = MIMEImage(temp.read(), _subtype=sub_type)
                temp.close()

            elif main_type == 'audio':
                print("audio attached")
                temp = open(attached_file, 'rb')
                attachment = MIMEAudio(temp.read(), _subtype=sub_type)
                temp.close()            

            elif main_type == 'application' and sub_type == 'pdf':   
                temp = open(attached_file, 'rb')
                attachment = MIMEApplication(temp.read(), _subtype=sub_type)
                temp.close()

            else:                              
                attachment = MIMEBase(main_type, sub_type)
                temp = open(attached_file, 'rb')
                attachment.set_payload(temp.read())
                encoders.encode_base64(attachment)
                temp.close()

            filename = attached_file.name
            attachment.add_header('Content-Disposition', 'attachment', filename=filename) # name preview in email
            self.attach(attachment)


class GoogleClient:
    """
    Class representing a Google client

    Author: Indrajit Ghosh

    Date: Aug 28, 2022

    N.B: To create a `GoogleClient` instance a `credentials.json` file is needed.
        This file is provided by Google. A `token.json` will be created after the
        first run.

    Parameters:
    -----------
        `client`:str; it can be 'drive' or 'gmail'
        `authorized_user_file`: optional; this is the path to authorized user file : `token.json`
        `client_secret_file`: optional; if not given then there should be a `credentials.json` file
                                in this script's directory
    """

    # If modifying these scopes, delete the file token.json.
    SCOPES = [
        'https://www.googleapis.com/auth/drive',
        'https://mail.google.com/',
        'https://www.googleapis.com/auth/spreadsheets'
    ]

    AUTHORIZED_USER_FILE = Path(__file__).resolve().parent / 'token.json'
    CLIENT_SECRET_FILE = Path(__file__).resolve().parent / 'credentials.json'

    def __init__(
        self,
        client:str='drive',
        authorized_user_file=AUTHORIZED_USER_FILE,
        client_secret_file=CLIENT_SECRET_FILE
    ):
    
        # Getting credentials
        self.credentials = self.get_creds(authorized_user_file=authorized_user_file, client_secret_file=client_secret_file)

        if client.lower() == 'drive':
            # create drive api client
            self.service = build('drive', 'v3', credentials=self.credentials)
        elif client.lower() == 'gmail':
            # create gmail api client
            self.service = build('gmail', 'v1', credentials=self.credentials)
        elif client.lower() == 'sheet':
            # create google sheet api client
            self.service = build('sheets', 'v4', credentials=self.credentials)


    def get_creds(self, authorized_user_file, client_secret_file):
        """
        Get Google credential

        Parameter
        ---------
            `authorized_user_file`: `token.json`
            `client_secret_file`: `credentials.json`
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(authorized_user_file):
            creds = Credentials.from_authorized_user_file(authorized_user_file, self.SCOPES)

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    client_secret_file, self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(authorized_user_file, 'w') as token:
                token.write(creds.to_json())

        return creds


class GoogleServiceAccountClient:
    """
    Service Account client
    """
    SCOPES = [
        'https://www.googleapis.com/auth/drive',
        'https://mail.google.com/',
        'https://www.googleapis.com/auth/spreadsheets'
    ]

    SERVICE_ACC_KEYS_FILE = Path(__file__).resolve().parent / 'keys.json'

    def __init__(self, client='sheet'):
        
        creds = None
        try:
            creds = service_account.Credentials.from_service_account_file(self.SERVICE_ACC_KEYS_FILE, scopes=self.SCOPES)

            # Getting credentials
            self.credentials = creds

            if client.lower() == 'drive':
                # create drive api client
                self.service = build('drive', 'v3', credentials=self.credentials)
            elif client.lower() == 'gmail':
                # create gmail api client
                self.service = build('gmail', 'v1', credentials=self.credentials)
            elif client.lower() == 'sheet':
                # create google sheet api client
                self.service = build('sheets', 'v4', credentials=self.credentials)


        except HttpError as err:
            print(f'ERROR: {err}')


class GoogleDrive:
    """
    Class representing a google drive

    Author: Indrajit Ghosh

    Date: Aug 28, 2022
    """

    MIMETYPES = {
    "gdoc": 'application/vnd.google-apps.document',
    "gsheet": 'application/vnd.google-apps.spreadsheet',
    "gslide": 'application/vnd.google-apps.presentation',
    "gform": 'application/vnd.google-apps.form'
    }

    def __init__(self, **kwargs):

        # create drive api client
        drive = GoogleClient(client='drive', **kwargs)
        self.service = drive.service


    def get_filelist(self, directory_id=None, N:int=None):
        """
        It returns a `list` of the last `N` many files of the directory with id `directory_id`

        Example
        -------
            [
                {"id": '1ftoucUbxz10c1zoU1mDuqMr2NMIxhmbi', 'name': 'spam1', 'mimeType': 'application/pdf'},
                {"id": '2934jlhdsd10c1zoU1mDuqMr2NMI3jksd', 'name': 'spam2', 'mimeType': 'application/pdf'},
            ]
        """

        try:
            # create drive api client
            service = self.service

            query = f"'{directory_id}' in parents and trashed=false" if directory_id is not None else directory_id

            # request a list of first N files or folders with name and id from the API.
            res = service.files().list(
                q=query,
                pageSize=N, 
                fields="files(id, name, mimeType)",
                spaces="drive"
            ).execute()

            return res.get('files')

        except HttpError as err:
            print(f'ERROR: {err}')


    def get_file_id(self, filename:str, parent_cloud_dir_id=None):
        """
        Return the file id of a file in the drive
        """

        try:
            # create drive api client
            files_inside = self.get_filelist(directory_id=parent_cloud_dir_id)
            not_found = True

            for file in files_inside:
                if file['mimeType'] != 'application/vnd.google-apps.folder' and file['name'] == filename:
                    not_found = False
                    return file['id']
            if not_found:
                raise FileNotFoundError(f"No file found with the name `{filename}`")

        except HttpError as err:
            print(f'ERROR: {err}')


    def get_directory_id(self, dir_name:str):
        """
        Search for all directories with the name `dir_name` and if 
        found exactly one then returns its id
        """

        try:
            allFolders = self.service.files().list(
                q = f"mimeType = 'application/vnd.google-apps.folder' and name = '{dir_name}'",
                fields="files(id, name, mimeType)",
                spaces="drive"
            ).execute()

            if len(allFolders['files']) > 1:
                raise NotImplementedError(f"\nThere are multiple folders with name `{dir_name}`\n")

            elif allFolders['files'] == []:
                raise FileNotFoundError(f"\nNo folders found with the name `{dir_name}`\n")

            else:
                return allFolders['files'][0]['id']

        
        except HttpError as err:
            print(f'ERROR: {err}')


    def get_filename(self, file_id):
        """
        Accepts
        -------
            file or folder id

        Returns
        -------
            the filename
        """
        try:
            filename = self.service.files().get(fileId=file_id).execute()['name']

        except HttpError as err:
            print(f'ERROR: {err}')
            filename = None

        return filename


    def get_mimetype(self, file_id):
        """
        Accepts
        -------
            file or folder id

        Returns
        -------
            the mimetype
        """
        try:
            mimeType = self.service.files().get(fileId=file_id).execute()['mimeType']

        except HttpError as err:
            print(f'ERROR: {err}')
            mimeType = None

        return mimeType


    def upload_file(self, local_filepath, cloud_dir_id:str=None, print_status=True):
        """
        Returns:
        --------
            `str`: file_id of the uploaded file
        """

        local_filepath = Path(local_filepath)
        parents = [] if cloud_dir_id is None else [cloud_dir_id]

        try:
            # create drive api client
            service = self.service

            file_metadata = {
                "name": local_filepath.name,
                "parents": parents
            }

            file_media = MediaFileUpload(local_filepath)

            upload_file = service.files().create(body=file_metadata, media_body=file_media, fields='id').execute()

            if print_status:
                print(f"  Uploaded file: {local_filepath.name}\n  File id: {upload_file.get('id')}\n")

        except HttpError as error:
            print(f'An error occurred: {error}')
            upload_file = None

        return upload_file.get('id')


    @staticmethod
    def save_file(iofileObj, filename:str='untitled_file', download_dir=Path.cwd()):
        """
        Parameters
        ----------
            `iofileObj`: io.file
            `filename`: it could be Path() obj also
        """
        filePath = download_dir / filename
        with open(filePath, 'wb') as f:
            f.write(iofileObj)

    
    def download_file(self, real_file_id, download_dir=Path.cwd()):
        """
        Downloads a file
        
        Args:
        -----
            real_file_id: ID of the file to download

        Returns:
        -------
             IO object with location. which can be saved to the local dir by
                        ```
                            with open(filename, 'wb') as f:
                                f.write(img)
                        ```
        """

        try:
            # create drive api client
            service = self.service

            file_id = real_file_id
            file_name = self.get_filename(file_id=file_id)
            mimetype = self.get_mimetype(file_id=file_id)

            # pylint: disable=maybe-no-member
            if mimetype is None:
                request = service.files().get_media(fileId=file_id)
            elif mimetype in [self.MIMETYPES['gdoc'], self.MIMETYPES['gsheet'], self.MIMETYPES['gslide'], self.MIMETYPES['gform']]:

                if mimetype == self.MIMETYPES['gdoc']:
                    mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                elif mimetype == self.MIMETYPES['gsheet']:
                    mimetype = "text/csv" # you can change it to "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" also
                elif mimetype == self.MIMETYPES['gslide']:
                    mimetype = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                elif mimetype == self.MIMETYPES['gform']:
                    mimetype = "application/vnd.google-apps.script+json"

                request = service.files().export_media(fileId=file_id, mimeType=mimetype)
            else:
                request = service.files().get_media(fileId=file_id)


            file = io.BytesIO()
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f'Download {int(status.progress() * 100)}.')

        except HttpError as error:
            print(f'An error occurred: {error}')
            file = None

        iofileObj = file.getvalue()

        self.save_file(iofileObj=iofileObj, filename=file_name, download_dir=download_dir)

        return iofileObj


    def rename_file(self, file_id:str, new_filename:str, print_status=True):
        """
        Returns:
        --------
            None
        """

        try:
            service = self.service

            # File metadata
            file_metadata = {
                'name': new_filename,
                'mimeType': self.get_mimetype(file_id=file_id)
            }

            # Send the request to the API.
            service.files().update(fileId=file_id, body=file_metadata).execute()

            if print_status:
                print(f"\nFile renamed to `{new_filename}`\n")

        except HttpError as error:
            print(f'An error occurred: {error}')



    def create_folder(self, folder_name:str, parent_dir_id=None, parent_dir_name=None, print_status=True):
        """
        Parameters
        ----------
            `dir_name`: str; Name of the directory to be created

            `parent_dir`: the directory in which a new directory will be created
                If it is `None` then the new dir will be created at base location
                of the drive

        Returns
        -------
            `id` of the dir created
        """

        try:
            service = self.service

            if parent_dir_id is None and parent_dir_name is None:
                
                # dir_id = "Base Google Drive directory"
                parents = []

            elif parent_dir_id is not None and parent_dir_name is None:
                dir_id = parent_dir_id
                parents = [dir_id]

            elif parent_dir_id is None and parent_dir_name is not None:
                parents = [self.get_directory_id(dir_name=parent_dir_name)]

            else:
                dir_id = self.get_directory_id(dir_name=parent_dir_name)

                if dir_id == parent_dir_id:
                    parents = [dir_id]
                else:
                    raise NotImplementedError("Given cloud directory name and cloud directory id don't match!")


            new_dir_metadata = {
                        "name": folder_name,
                        "parents": parents,
                        "mimeType": 'application/vnd.google-apps.folder'
                    }
            new_cloud_dir = service.files().create(body=new_dir_metadata, fields="id").execute()

            if print_status:
                print("Directory Created successfully!")

            return new_cloud_dir.get('id')


        except HttpError as err:
            print(f'ERROR: {err}')


    def delete_file(self, file_id):
        """Permanently delete a file, skipping the trash.

        Args:
          service: Drive API service instance.
          file_id: ID of the file to delete.
        """
        try:
            service = self.service
            service.files().delete(fileId=file_id).execute()
            print("File deleted permanently!")

        except HttpError as err:
            print(f'ERROR: {err}')


    def create_blank_googlesheet(self, spreadsheet_name:str='Untitled', parent_dir_id:str=None):
        
        try:
            service = self.service

            if parent_dir_id is None:
                # dir_id = "Base Google Drive directory"
                parents = []

            else:
                dir_id = parent_dir_id
                parents = [dir_id]

            sheet_metadata = {
                "name": spreadsheet_name,
                "mimeType": 'application/vnd.google-apps.spreadsheet',
                "parents": parents
            }

            file = service.files().create(body=sheet_metadata, fields="id").execute()
            file_id = file.get('id')

        except HttpError as err:
            print(f'ERROR: {err}')
            file_id = None

        
        return file_id


class GmailMessage(EmailMessage):
    """
    A class representing GmailMessage

    Author: Indrajit Ghosh

    Date: Aug 28, 2022
    """

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        ## encode the message (the message should be in bytes)
        message_as_bytes = self.as_bytes() # the message should converted from string to bytes.
        message_as_base64 = base64.urlsafe_b64encode(message_as_bytes) #encode in base64 (printable letters coding)
        raw = message_as_base64.decode()  # need to JSON serializable

        self.gmail_message = {'raw': raw}


    def send(self):
        try:
            gmail_client = GoogleClient(client='gmail')
            service = gmail_client.service

            send_message = (service.users().messages().send
                            (userId="me", body=self.gmail_message).execute())
            print(f'Message Id: {send_message["id"]}')

        except HttpError as error:
            print(F'An error occurred: {error}')
            send_message = None


class Gmail:
    # TODO: Not really complete
    """
    Class representing a gmail acc

    Author: Indrajit Ghosh

    Date: Aug 28, 2022
    """

    def __init__(self):

        # create drive api client
        gmail = GoogleClient(client='gmail')
        self.service = gmail.service


    def get_messages(self):
        """
        Functions to print the mailbox
        """
        result = self.service.users().messages().list(userId='me').execute()

        messages = result.get('messages') # messages is a list of dictionaries where each dictionary contains a message id.

        return messages


    def get_email_message(self, message_id):
        """
        Returns:
        --------
            `subject`, `sender`, `body`
        """

        msg = self.service.users().messages().get(userId='me', id=message_id).execute()

        # Get value of 'payload' from dictionary `msg`
        payload = msg['payload']
        headers = payload['headers']

        # Look for Subject and Sender Email in the headers
        for d in headers:
            if d['name'] == 'Subject':
                subject = d['value']
            if d['name'] == 'From':
                sender = d['value']

        # The Body of the message is in Encrypted format. So, we have to decode it.
        # Get the data and decode it with base 64 decoder.
        parts = payload.get('parts')[0]
        data = parts['body']['data']
        data = data.replace("-","+").replace("_","/")
        decoded_data = base64.b64decode(data)


        # Now, the data obtained is in lxml. So, we will parse 
        # it with BeautifulSoup library
        soup = BeautifulSoup(decoded_data , "lxml")
        body = soup.body()

        return subject, sender, body


    @staticmethod
    def print_email_message(subject, sender, body=None):
        
        # Printing the subject, sender's email and message
        print()
        print('.'*100)
        print(" Subject: ", subject)
        print(" From: ", sender)

        if body is not None:
            print(" Message: ", body)

        print('.'*100)

        print('\n')


    def inbox(self, num_of_emails=5, print_body=False):

        messages = self.get_messages()

        for el in messages[:num_of_emails + 1]:
            msg_id = el['id']
            sub, sender, body = self.get_email_message(message_id=msg_id)

            if not print_body:
                body = None
            
            self.print_email_message(sub, sender, body)



class GoogleSheetClient(GoogleClient, GoogleServiceAccountClient):
    """
    A class representing a GoogleSheetClient

    Author: Indrajit Ghosh
    Date: Nov 15, 2022

    API_URL: "https://developers.google.com/sheets/api/reference/rest"
    """

    def __init__(self, service_acc=False):

        if not service_acc:
            GoogleClient.__init__(self, client='sheet')
        else:
            GoogleServiceAccountClient.__init__(self, client='sheet')


    def create_spreadsheet(
        self,
        spreadsheet_title:str="Untitled",
        num_of_rows:int=1000,
        num_of_cols:int=26,
        sheet1_title:str="Sheet 1"
    ):
        """
        Creates a new spreadsheet

        Accepts:
        --------

        """

        spreadsheet_body = {
            "properties": {
                "title": spreadsheet_title,
            },
            "sheets": [
                {
                    "properties": {
                        "title": sheet1_title,
                        "gridProperties": {
                            "rowCount": num_of_rows,
                            "columnCount": num_of_cols
                        },
                        "sheetType": 'GRID',
                    }
                }
            ]
        }

        gsheet = self.service.spreadsheets().create(body=spreadsheet_body).execute()
        
        pprint(gsheet)


    def open_by_gspreadsheet_id(self, spreadsheet_id:str):
        """
        Open a Google spreadsheet by id

        Returns:
        --------
            `GoogleSheet` object
        """

        return GoogleSheet(client=self, spreadsheet_id=spreadsheet_id)


class GoogleSheet:
    """
    A class representing a GoogleSpreadSheet

    Author: Indrajit Ghosh
    Date: Nov 15, 2022
    """

    def __init__(self, client:GoogleSheetClient, spreadsheet_id:str):

        self.client = client
        self.id = spreadsheet_id
        self.service = client.service

        self.set_properties()


    def __repr__(self):
        return "<{} {} id:{}>".format(
            self.__class__.__name__,
            repr(self.title),
            self.id,
        )

    def set_properties(self):
        """Sets up basic properties"""
        # The following gives a `dict` with keys: dict_keys(['spreadsheetId', 'properties', 'sheets', 'spreadsheetUrl'])
        # using this `dict` we can create GoogleSheet object
        self.spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.id).execute()
        self._properties = self.spreadsheet['properties'] # dict
        self._sheets = self.spreadsheet['sheets'] # list

    @property
    def title(self):
        """GoogleSpreadsheet Title"""
        return self._properties["title"]

    @property
    def url(self):
        """Spreadsheet URL."""
        return self.spreadsheet['spreadsheetUrl']
    
    @property
    def sheet1(self):
        """Shortcut property for getting the first worksheet."""
        return self.get_worksheet(0)

    @property
    def timezone(self):
        """Spreadsheet timeZone"""
        return self._properties["timeZone"]

    @property
    def locale(self):
        """Spreadsheet locale"""
        return self._properties["locale"]

    def worksheets(self):
        """
        Worksheets of that GoogleSpreadsheet

        Returns:
        --------
            `list[Worksheet(), ... , Worksheet()]`
        """
        wkshts = self.spreadsheet['sheets'] # `list` of `dict`

        return [Worksheet(self, wks['properties']) for wks in wkshts]


    def get_worksheet(self, index):
        """
        Returns a worksheet with specified `index`.

        :param index: An index of a worksheet. Indexes start from zero.
        :type index: int

        Returns:
        --------
            `Worksheet()`

        Example. To get third worksheet of a spreadsheet:
        >>> sht = client.open('My fancy spreadsheet')
        >>> worksheet = sht.get_worksheet(2)
        """

        try:
            return self.worksheets()[index]

        except (KeyError, IndexError):
            raise GoogleSheetError("index {} not found".format(index))


    def get_worksheet_by_title(self, sheet_title:str):
        """
        Get worksheet by it's title

        Accepts:
        --------
            `sheet_title`: str

        Returns:
        --------
            `Worksheet()`
        """
        for sheet in self._sheets:
            if sheet['properties']['title'] == sheet_title:
                sheet_index = sheet['properties']['index']
                return self.get_worksheet(index=sheet_index)

        raise WorksheetNotFound(f"No worksheet with the title `{sheet_title}` found!")

    
    def has_worksheet_with_given_title(self, title:str):
        """
        Checks whether the GoogleSheet has a `worksheet` by the given title.

        Parameter:
        ----------
            `title`: `str`
        
        Returns:
        --------
            `Bool`
        """
        available_worksheet_titles = [sheet['properties']['title'] for sheet in self._sheets]
        return title in available_worksheet_titles


    def get(self, range_, majorDimension='ROWS'):
        """
        Gets the spreadsheet values at the given `range_`

        Parameter:
        ----------
            `range_`: str in A1 notation

        Returns:
        --------
            `list[list[], ... , list[]]`

        Example:
        --------
            >>> spreadsheet = GoogleSheet()
            >>> spreadsheet.get(range_='Sheet 1!A1:B1')
                [['Timestamp', 'Price (INR)']]
            >>> spreadsheet.get(range_='Sheet 1!A1')
                [['Timestamp']]
            
        """
        response = self.service.spreadsheets().values().get(
            spreadsheetId=self.id,
            range=range_,
            majorDimension=majorDimension
        ).execute()

        return response['values']


    def update(self, range_, values:list):
        """
        Updates the spreadsheet at the given `range_`

        Parameters:
        ----------
            `range_`: str in A1 notation
            `values`: list[list[], ... ,list[]]

        Example:
        --------
            >>> spreadsheet = GoogleSheet()
            >>> spreadsheet.update(range_='Sheet 1!A1:B1', [[1, 3]])
        """

        value_range_body = {
            "range": range_,
            "majorDimension": 'ROWS',
            "values": values
        }

        response = self.service.spreadsheets().values().update(
            spreadsheetId=self.id,
            range=range_,
            valueInputOption='USER_ENTERED',
            body=value_range_body
        ).execute()

        pprint(response)


    def batch_requests_update(self, list_of_requests:list):
        """
        This function can handle one or more Requests for 
        updating the GoogleSheet.

        Parameter:
        ----------
            `list_of_requests`: `list`
                Structure of this:
                    [
                        {
                            object (Request)
                        },
                        {
                            object (Request)
                        }
                    ]
        """
        batch_update_spreadsheet_request_body = {
            "requests": list_of_requests
        }

        request = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.id,
            body=batch_update_spreadsheet_request_body
        )
        response = request.execute()

        self.set_properties()

        pprint(response)


    def add_worksheet(self, title:str, index:int=None, worksheet_type:str=None, grid_properties:dict=None):
        """
        Add worksheet into the GoogleSheet

        Reference:
        ----------
            "https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#addsheetrequest"

        Parameters:
        -----------
            `worksheet_type`: `str` # URL: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#sheettype

                    SHEET_TYPE_UNSPECIFIED	    Default value, do not use.
                    GRID	                    The sheet is a grid.
                    OBJECT	                    The sheet has no grid and instead has an object like a chart or image.
                    DATA_SOURCE	                The sheet connects with an external DataSource and shows the preview of data.


            `grid_properties`: `dict` # URL: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#gridproperties
                    Example: {
                                "rowCount": integer,
                                "columnCount": integer,
                                "frozenRowCount": integer,
                                "frozenColumnCount": integer,
                                "hideGridlines": boolean,
                                "rowGroupControlAfter": boolean,
                                "columnGroupControlAfter": boolean
                            } 
        """
        index = len(self._sheets) if index is None else index
        worksheet_type = 'GRID' if worksheet_type is None else worksheet_type.upper() 
        grid_properties = {"rowCount":1000, "columnCount":26} if grid_properties is None else grid_properties

        add_worksheet_request = {
            "addSheet": {
                "properties": {
                    "title": title,
                    "index": index,
                    "sheetType": worksheet_type,
                    "gridProperties": grid_properties
                }
            }
        }

        self.batch_requests_update(list_of_requests=[add_worksheet_request])


    def _delete_worksheet_by_id(self, worksheet_id:int):
        """
        Deletes the worksheet from the GoogleSheet using its sheetId
        """

        del_worksheet_request = {
            "deleteSheet": {
                "sheetId": worksheet_id
            }
        }

        self.batch_requests_update(list_of_requests=[del_worksheet_request])


    def delete_worksheet(self, title:str):
        """
        Deletes the worksheet
        """
        noSheet = True
        for sheet in self._sheets:
            if sheet['properties']['title'] == title:
                self._delete_worksheet_by_id(worksheet_id=sheet['properties']['sheetId'])
                noSheet = False
                break

        if noSheet:
            raise WorksheetNotFound(f"No worksheet found with the title `{title}`.")


    @classmethod
    def create_spreadsheet(cls, spreadsheet_title:str=None, **kwargs):
        """
        Creates a new spreadsheet

        Accepts:
        --------
            `str`: spreadsheet_title - optional

        """
        client = GoogleSheetClient(**kwargs)

        if not spreadsheet_title:
            gsheet = client.service.spreadsheets().create().execute()
        
        else:
            spreadsheet_body = {
                'properties':{
                    'title': spreadsheet_title,
                    'locale': 'en_US',
                    'timeZone': 'Asia/Calcutta'
                },
                'sheets':[
                    {
                        'properties':{
                            'title': 'Sheet 1'
                        }
                    }
                ]
            }

            gsheet = client.service.spreadsheets().create(
                body=spreadsheet_body
            ).execute()
        
        return GoogleSheet(client=client, spreadsheet_id=gsheet['spreadsheetId'])


class Worksheet:
    """
    The class that represents a single sheet in a spreadsheet
    (aka "worksheet").

    Author: Indrajit Ghosh
    Date: Nov 16, 2022
    """
    def __init__(self, spreadsheet:GoogleSheet, properties):
        
        self._properties = properties
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client


    def __repr__(self):
        return "<{} {} id:{}>".format(
            self.__class__.__name__,
            repr(self.title),
            self.id,
        )

    
    @property
    def id(self):
        """Worksheet ID."""
        return self._properties["sheetId"]

    @property
    def title(self):
        """Worksheet title."""
        return self._properties["title"]

    @property
    def index(self):
        """Worksheet index."""
        return self._properties["index"]

    @property
    def row_count(self):
        """Number of rows."""
        return self._properties["gridProperties"]["rowCount"]

    @property
    def col_count(self):
        """Number of columns.
        .. warning::
           This value is fetched when opening the worksheet.
           This is not dynamically updated when adding columns, yet.
        """
        return self._properties["gridProperties"]["columnCount"]

    def get_cell(self, coordinate:str):
        """
        Gets the cell value

        Parameters:
        -----------
            `coordinate`: `str` In A1 notation; e.g. 'A3', 'C5' etc
        """
        rng = self.title + "!" + coordinate
        return self.get(range_=rng)[0][0]

    def update_cell(self, coordinate:str, value):
        """
        Updates the cell value

        Parameters:
        -----------
            `coordinate`: `str` In A1 notation; e.g. 'A3', 'C5' etc
            `value`: `Any`
        """
        return self.update(range_=coordinate, values=[[value]])


    def get(self, range_=None, majorDimension:str='ROWS'):
        """
        Get values of a range or cell in A1 notation

        Parameters:
        -----------
            `ranges`: ranges in A1 notation
            `majorDimension`: The major dimension that results should use.
                        For example, if the spreadsheet data on Sheet1 is: A1=1,B1=2,A2=3,B2=4, then 
                        requesting range=Sheet1!A1:B2?majorDimension=ROWS returns [[1,2],[3,4]], whereas 
                        requesting range=Sheet1!A1:B2?majorDimension=COLUMNS returns [[1,3],[2,4]].


        Returns:
        --------
            `list[list[], ... , list[]]`


        Examples:
        ---------
             # Return all values from the sheet
             worksheet.get()

             # Return value of 'A1' cell
             >>> worksheet.get('A1')
                [['Timestamp']]

             # Return values of 'A1:B1' range
             >>> worksheet.get('A1:B1')
                [['Timestamp', 'Price (INR)']]
        """
        if range_ is None:
            range_ = self.title
        elif isinstance(range_, str):
            if len(range_) == 2:
                range_ = self.title + "!" + range_
            elif ":" in range_:
                range_ = self.title + "!" + range_

        response = self.spreadsheet.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet.id,
            range=range_,
            majorDimension=majorDimension
        ).execute()

        return response['values']


    def update(self, range_, values:list):
        """
        Updates the worksheet at the given `range_`

        Parameters:
        ----------
            `range_`: str in A1 notation
            `values`: list[list[], ... ,list[]]

        Example:
        --------
            >>> wks = Worksheet()
            >>> wks.update(range_='A1:B1', [[1, 3]])
        """
        range_ = self.title + '!' + range_

        self.spreadsheet.update(range_=range_, values=values)


    def insert_rows(self, start_index:int=1, number_of_rows:int=1):
        """
        Inserts `number_of_rows` many rows after the row with index `start_index`
        """
        self.insert(dimension='ROWS', start_index=start_index, end_index=start_index + number_of_rows)

    def insert_columns(self, start_index:int=0, number_of_cols:int=1):
        """
        Inserts `number_of_cols` many cols after the col with index `start_index`
        """
        self.insert(dimension='COLUMNS', start_index=start_index, end_index=start_index + number_of_cols)


    def insert(self, dimension:str='ROWS', start_index:int=0, end_index:int=1):
        """
        Inserts `rows`(or cols) at a particular index inside the worksheet.

        Suppose `dimension`= 'ROWS', then this function insert (end_index - start_index)
        many rows after the row with index `start_index` and similar for 'COLUMNS' case.

        Parameters:
        -----------
            `dimension`:str; it could be 'ROWS' or 'COLUMNS'
        """

        batch_update_spreadsheet_request_body = {
            "requests": [
                {
                    'insertDimension': {
                        "range": {
                            "sheetId": self.id,
                            "dimension": dimension,
                            "startIndex": start_index,
                            "endIndex": end_index
                        }
                    }
                }
            ]
        }
    
        # Inserting
        response = self.spreadsheet.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet.id,
            body=batch_update_spreadsheet_request_body
        ).execute()

        pprint(response)



def main():
    print('Classes required for Google APIs')

if __name__ == '__main__':
    main()