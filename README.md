#Google Sheets Automation Script
##Project Overview
This project is a Google Apps Script designed to automate the creation and management of an "Index" tab and a "Cover" tab within a Google Spreadsheet. It performs various tasks such as adding hyperlinks, filtering tab information, deleting unused rows, and applying conditional formatting and design. This script enhances efficiency for managing multiple tabs in Google Sheets by generating an easy-to-navigate index and a summary cover.

##Features
Index Tab Automation:

Automatically creates an "Index" tab if it doesn't exist.
Populates the "Index" tab with a list of all sheet names (excluding the "Index" and "Cover" tabs) and their hyperlinks.
Adds a dropdown in the "Obsolete/Current" column for easy status tracking.
Retains any existing values in the "About" column without overwriting them.
Deletes blank rows and unused columns in the "Index" tab.
Cover Tab Automation:

Creates a "Cover" tab if it doesn't exist.
In the "Cover" tab, populates columns with 25 records per column from the "Index" tab, filtering only the "Current" status.
Uses dynamic Google Sheets formulas such as IFERROR and FILTER for efficient data filtering.
Design & Formatting:

Applies custom formatting: black background, white text, bold, and center-aligned headers for both the "Index" and "Cover" tabs.
Automatically resizes columns in both tabs to fit content.
##Installation
###Prerequisites
A Google account with access to Google Sheets.
Basic understanding of Google Apps Script (optional but useful).
###Setup Instructions
Open your Google Spreadsheet.
Click on Extensions > Apps Script to open the script editor.
Copy the latest version of the script from this repository and paste it into the script editor.
Save the script with a meaningful name (e.g., Sheet Automation Script).
Go to the Triggers section (Triggers icon on the left panel), and create a new trigger:
Select the function onOpen.
Choose event type: From spreadsheet > On open.
Save the trigger and close the Apps Script editor.
##Usage
Once the script is installed and saved, it will run automatically each time the spreadsheet is opened.
##The script will:
Create and update the "Index" and "Cover" tabs dynamically.
Add hyperlinks to the sheet names in the "Index" tab.
Filter and display sheet information in the "Cover" tab based on the status in the "Index" tab.
If any changes are made to the sheet names, the "Index" and "Cover" tabs will update accordingly.
Example
The script automatically filters the first 25 "Current" status sheets from the "Index" tab and displays them in the first column of the "Cover" tab, the next 25 in the second column, and so on. It formats both the "Index" and "Cover" tabs with a black background and white, bold text for easier navigation and presentation.

##Best Practices
Version Control: Keep track of changes to the script by committing meaningful updates to this repository regularly.
Performance Optimization: The script is designed to minimize unnecessary recalculations and updates. Ensure that your spreadsheet doesn't have excessive tabs for optimal performance.
Error Handling: The use of IFERROR ensures that empty or error-prone data entries are handled gracefully without breaking the spreadsheet functionality.
Reusability: The script is modular and can be easily adapted for similar spreadsheet management tasks in other projects.
##License
This project is licensed under the MIT License – see the LICENSE file for details.
