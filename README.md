# Work Time Tracker
A simple time-tracking application built with Kivy for desktop use. This app allows users to clock in and out of tasks, save time logs to a CSV file, and generate a weekly summary in PDF format. Additionally, it integrates with Microsoft SharePoint to upload task data, making it suitable for team environments that use SharePoint for project management.

# Features
Clock In/Out: Track the start and end of tasks with a simple click.
Task Description: Enter a task description before clocking in.
Generate Weekly PDF Summary: Create a PDF summarizing hours worked, task descriptions, and clock-in/out times.
SharePoint Integration: Upload task data (task description, clock-in/out time, etc.) directly to a SharePoint list.
Settings: Configure SharePoint client settings directly within the app.

# Requirements
Python 3.x
Kivy: GUI framework for building the app.
MSAL: Microsoft Authentication Library for handling token acquisition and authentication with SharePoint.
ReportLab: Used for generating the PDF reports.
Requests: HTTP library for SharePoint API interaction.

# To install the required dependencies, run the following command:
bash
Copy code
pip install kivy msal requests reportlab

# Installation
Clone the repository:
bash
Copy code
git clone https://github.com/username/repository-name.git
cd repository-name
Install dependencies:

bash
Copy code
pip install -r requirements.txt
Run the application:

bash
Copy code
python app.py
Configuration

# SharePoint Settings
# Before using the SharePoint integration, you need to configure your SharePoint client settings. This can be done directly from the app’s Settings screen. You'll need:

Client ID
Authority URL
SharePoint Site URL
List Name (the list where tasks will be logged)
Once entered, the app will save these settings to a settings.json file, which is used for interacting with SharePoint.

# Authentication
The app uses Microsoft Authentication (MSAL) to authenticate with SharePoint. If no valid access token is found, the app will prompt the user to sign in.

# Usage
Clock In: Enter your ID, Department and, task description and press "Clock In".
Clock Out: When finished, press "Clock Out" to log the time spent on the task.

**Generate PDF:**  After clocking out, you can generate a weekly summary PDF by clicking the "Generate Weekly Summary" button.
SharePoint Integration: When clocking out, the task details will be uploaded to the configured SharePoint list.

# File Structure
bash
Copy code
/work-time-tracker
    /app.py                  # Main application code
    /settings.json           # Stores SharePoint settings
    /time_log.csv            # Stores logged clock-in/clock-out data
    /weekly_summary.pdf      # Generated PDF report
    /requirements.txt        # Python dependencies

    Contact
For questions or support, please contact:

Solomon N.
Email: smartzgh@gmail.com

# License
 GNU GENERAL PUBLIC LICENSE
 Version 3, 29 June 2007

