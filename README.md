# External Drives Documentation Tool
A Python program to scan external drives connected to your machine, collect certain properties, and store the information in a Google Spreadsheet.  
It is designed for a pool of external drives shared between multiple users. The documentation can get updated from different machines, even different operating systems.
Google Spreadsheet conveniently features a version history with back-ups, sorting options, search functions and (color) formatting.
However, there is still an option to use a JSON-File as documentation, e.g. if you want to connect to your own system.

---

## Features

- Scan connected external drives and collect the following properties:
- Properties of the drives:
  - Name
  - Total storage size
  - Free storage
- Properties of the First-level directories (projects):
  - Name
  - Size
  - Date (derived from the first 10 characters of the folder name, if it matches the YYYY-MM-DD format)

- All collected information get saved/updated in the documentation
- Blacklist excludes user-set drive- and directory-names

### Google Spreadsheet Documentation
- Documentation is maintained in a Spreadsheet using a Bot (Google Service Account)
- Drives are color-coded by automatically created Conditional Formatting Rules

### JSON Documentation
- Alternatively, the documentation can be saved as a JSON Document (its Structure can be found in the Appendix of this document)

---

##  Installation

### Google Spreadsheet and Service Account Credentials Set Up
- The Spreadsheet for your documentation should copy the Form of this Spreadsheet https://docs.google.com/spreadsheets/d/1JO8Urt_x54ECY-hZKj2Vi0Oxf6CVQkEFgwY1gUro0go/edit?usp=sharing. You can simply copy it (There is a guide in the Appendix).
- You can name the Spreadsheet however you like, but you would have to change the corresponding constant. It is easiest to just mirror the name ("external drives documentation").
- You have to create a Google Service Account:
  - Follow the steps 1-6 and of this guide: https://docs.gspread.org/en/latest/oauth2.html#for-bots-using-service-account
  - Rename the file to service_account.json and put it in the project-folder (the Installer will move service_account.json to its right place)
  - Lastly, to permit the bot access to our Spreadsheet you have to invite it.
    - To do so, open the JSON-File. You can use TextEdit (mac) or Editor (windows); select it by right-click on service_account.json > open with... . 
    - Copy the client_email, which should look something like this `473000000000-yoursisdifferent@developer.gserviceaccount.com`
    - Finally, you have to go to your Spreadsheet, click Share on the top right and insert the email in the field and invite your bot.

### Actual Installation
- Clone Repository and set up python environment  

macOS
```
git clone https://github.com/LenardJSchnakenbeck/create_drives_documentation_eventtec
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```
Windows
```
git clone https://github.com/LenardJSchnakenbeck/create_drives_documentation_eventtec
python -m venv venv
venv/Script/activate
pip install -r requirements.txt
```


## Contact / Author
https://www.linkedin.com/in/lenard-schnakenbeck/

## Appendix
### Copy a View-Only Google Sheet
1. Open the Google Sheet: Start by opening the Google Sheet you want to copy. Ensure you are logged into the correct Google account that has access to the sheet.
2. Go to File: In the menu bar at the top, click on File.
3. Make a Copy: From the dropdown menu, select Make a copy. A dialog box will appear prompting you to name the new copy and choose its destination folder in your Google Drive.
4. Choose a Name and Location: Name your copy something distinct, so you can easily differentiate it from the original. Select a folder in your Google Drive where you’d like to save it. This step is crucial for organizing your documents efficiently.
5. Click OK: Once you’re done naming and choosing a location, click OK. Voilà! You now have a copy of your Google Sheet that you can modify independently of the original.
(Source: https://www.thebricks.com/resources/guide-how-to-make-a-copy-of-google-sheets-that-is-view-only, 10.11.2025 18:00)
### Exact requirements for the Google Sheet
- Spreadsheet name: "external drives documentation"
- min. 2 Worksheets
- the second Worksheet should have the columns "blacklist-drives" and "blacklist-projects"
- shared to the E-Mail of the Google Service Account
### Structure of the JSON-Documentation
{  
  "drive 1": {  
  "name": "drive 1",  
    "total-storage (Gigabyte)": 112.805,  
    "free-storage (Gigabyte)": 11.713,  
    "projects": [  
      {  
        "name": "01-01-2025 Project A",  
        "size (Gigabyte)": 10.023,  
        "date": "01-01-2025"  
      },  
      {  
        "name": "02-02-2025 Project B",  
        "size (Gigabyte)": 125.567,  
        "date": "02-02-2025"  
      }  
    ]  
  },  
  "drive 2": {  
    "name": "drive 2",  
    "total-storage": 112.805,  
    "free-storage": 11.714,  
    "projects": [...]  
  },  
  ...  
}  
