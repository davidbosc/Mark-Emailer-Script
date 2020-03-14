# Mark-Emailer-Script
A set of scripts to facilitate emailing marks out for TA responsibilities, by creating emails in outlook with bodies populated from Google Sheets, or by creating txt files with HTML <br> tags to copy/paste into Nexus.

## Requirements
- Python 3.7
- A python package manager of your choosing (pip, anaconda, etc.)
- Microsoft Outlook
- Google Sheets API

### Dependencies
- \_\_future\_\_
- pickle
- os.path
- googleapiclient.discovery
- google_auth_oauthlib.flow
- google.auth.transport.requests
- csv
- win32com.client
- time

## How to Use

### Preparations

Follow [this guide](https://developers.google.com/sheets/api/quickstart/python) to obtain a google sheets API key, saving your credentials.json in the cloned repo's directory.  Also ensure that you have a token.pickle file generated and saved in the same location.

The *send_grades.py* script uses .csv files with key value pairs to populate class details.  Namely, the _google_sheet_ids.csv_ file should contain the name of a google sheet, and the sheet's unique string in a URL.  For example, if class1 had lab grades in a google sheet with the URL https://docs.google.com/spreadsheets/d/abcde12345/edit#gid=0, we'd store this as follows in the CSV: 
> class1L, abcde12345

Each entry in google_sheets_ids.csv should have a matching csv file that contains names and email addresses.  The names could either be a single student or a group name, but if a group is used, email addresses must be seperated by semi-colons such as:
>team1, email1@gmail.com;email2@gmail.com

Here is an example of the file structure.  All csv files are class lists, except google_sheets_ids.

![img](https://i.imgur.com/eOD5VCJ.png)

### SummaryOfRubricSpreadsheetFunctions.gs

SummaryOfRubricSpreadsheetFunctions.gs is a google script that adds the functions to google sheets to create a summary tab for a sheet.  The functions are:

__sheetName__

Takes an integer as a parameter that maps a sheet tab to the tab's name.  Sheets should be named after the student, so the csv file can map their name to their email.

> @param idx   Index of desired sheet.  Indexing starts at 0 to # of sheets -1, but if your first sheet is a summary, your starting index should always be 1.

> @return      Sheet name mapped from idx

__getMarkFromSheet__

Takes tab, row and column indexes, and retreives the final mark.

> @param idx  Index of desired sheet.  Indexing starts at 0 to # of sheets -1, but if your first sheet is a summary, your starting index should always be 1.

> @param row  Row where the final mark is stored in a spreadsheet.  Should be constant accross all sheets in the spreadsheet.

> @param col  Column where the final mark is stored in a spreadsheet.  Should be constant accross all sheets in the spreadsheet.

> @return     The value stored in the desired tab and specific row and column

__getFeedbackFromSheet__

Defines a range of cells to check for content.  All range of cells with content will be concatenated into one block string, breaking lines for each cell and will add a <br> for email formatting. 

> @param questionRowStart   Starting row for feedback cells.

> @param questionRowEnd     Ending row for feedback cells.

> @param questionColStart   Starting column for feedback cells.

> @param questionColEnd     Ending column for feedback cells.

> @return     All non-empty cells of feedback concatenated into a string block.  New lines are seperated by <br> tags for email formatting.

To use these functions in a google sheet, open your desired sheet and follow **Tools > Script Editor**.  Then, copy and paste the code content into the code editor and save the project.

### send_grades.py

Once you have a sheet populated you can run this script.  The script's workflow is:

- elect a mode to execute this script:
  - currently supports sending grades via email, or creating txt files to copy/paste into nexus
- (Email only) Prompt for preview before sending: 
  - Yes will wait for you to send the email before creating the next one.
  - No will create them all at once.
- Choose a spreadsheet:
  - These are the keys from the google_sheet_ids.csv.  It is a good idea to have your classlist CSVs to be named the same as the keys here so it's obvious which is which
- Enter a spreadsheet tab name:
  - This is where data is pulled from.  Likely a summary sheet, so it's a good idea to name that sheet after the assignment as this will be used in the email body to say that a mark for "Assignment x" is ..., etc (Email only).
- Enter weight (default is /10)
  - Weight of the assignement/lab/project.
  - Entering nothing will default the value to 10, as expected.
