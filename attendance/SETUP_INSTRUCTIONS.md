# Badminton Club Attendance System Setup Instructions

## Prerequisites
1. Google Account with access to Google Sheets and Google Apps Script
2. Your Google Sheet ID: `1yoqVoIMuFlcgm0gvNmu7fOXwyt1tHlyi-GVXn5xMweI`

## Setup Steps

### Step 1: Prepare Google Sheets
1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1yoqVoIMuFlcgm0gvNmu7fOXwyt1tHlyi-GVXn5xMweI/edit
2. The system will automatically create the required sheets:
   - **Attendance** (with columns: Date, Name, Email, Availability, Reason, Attended, Warning Count, Missed Days, AutoRemove)
   - **Players** (with columns: Name, Email)
   - **Hosts** (with columns: Name, Password, Role, IsAdmin, Email)

### Step 2: Set Up Google Apps Script
1. Go to [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Delete the default `Code.gs` content
4. Copy and paste the content from `Code.gs` file provided
5. Save the project with a meaningful name (e.g., "Badminton Attendance System")

### Step 3: Initialize the System
1. In Google Apps Script, run the `initializeSheets()` function once:
   - Click on the function dropdown and select `initializeSheets`
   - Click the "Run" button
   - Grant necessary permissions when prompted
2. Run the `setupTriggers()` function once:
   - Select `setupTriggers` from the dropdown
   - Click "Run"
   - This sets up automatic daily tasks and scheduled triggers:
     - **dailySetup**: 8:00 AM IST
     - **checkWarningsAndInactivity**: 11:00 PM IST
     - **sendAttendanceReminder**: 9:00 PM IST
     - **sendAvailabilitySummary**: 10:30 PM IST
     - **sendCelebrationIfFourAvailable**: every 10 minutes (all day)
     - **autoSetAttendedToAvailability**: 12:01 AM IST

### Step 4: Deploy as Web App
1. In Google Apps Script, click "Deploy" > "New deployment"
2. Choose "Web app" as the type
3. Set the following configurations:
   - Description: "Badminton Attendance System"
   - Execute as: "Me"
   - Who has access: "Anyone" (or "Anyone with Google account" for more security)
4. Click "Deploy"
5. Copy the Web App URL (it will look like: `https://script.google.com/macros/s/[SCRIPT_ID]/exec`)

### Step 5: Update HTML File
1. Open `index.html`
2. Find this line near the top of the JavaScript section:
   ```javascript
   const SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
3. Replace `YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with your actual Web App URL from Step 4

### Step 6: Host the HTML File
You can host the HTML file in several ways:
1. **GitHub Pages** (Free):
   - Create a GitHub repository
   - Upload `index.html`
   - Enable GitHub Pages in repository settings
   
2. **Google Sites** (Free):
   - Create a new Google Site
   - Embed the HTML using an HTML embed component
   
3. **Any web hosting service** (Netlify, Vercel, etc.)

### Step 7: Initial Setup
1. Open your hosted HTML page
2. The system will detect that it's empty and show "Initial Setup"
3. Enter your name, email, and create an admin password
4. Click "Setup as Main Host"
5. You are now the main administrator!

## System Features

### Automatic Daily Tasks & Triggers
The system automatically:
- Creates daily attendance entries at 8:00 AM IST
- Sends attendance reminder at 9:00 PM IST
- Sends availability summary at 10:30 PM IST
- Checks for warnings and inactivity at 11:00 PM IST
- Sets attended status automatically at 12:01 AM IST
- Checks for 4 available players every 10 minutes

### Player Workflow
1. Players submit availability by 10:30 PM the previous day (cutoff enforced)
2. They select "Yes" or "No" for tomorrow's session
3. Can provide optional reason if selecting "No"
4. Must enter both name and email (matching Players sheet)

### Host Workflow
1. Hosts can update actual attendance before 10:00 AM IST
2. Must authenticate with name, email, and password
3. System automatically applies warning logic
4. Host can only update "Attended" for today

### Admin Management
1. Add/remove players (name and email required)
2. Add/remove hosts and co-hosts (name, email, role, password required)
3. View all attendance records
4. Monitor player status and warnings

### Warning System
- "Yes" for availability but did not attend: warning added
- 5 warnings = automatic removal
- More than 15 missed days in a month = automatic removal
- 10 "No" without valid reason in a month = automatic removal
- All warnings and removals are tracked and displayed

## Security Features
- Host authentication required for attendance updates (name, email, password)
- Admin authentication required for system management
- Time-based restrictions (10:30 PM cutoff for players, 10:00 AM host update limit)
- Role-based access control
- Passwords are hashed client-side before sending to server

## Troubleshooting

### Common Issues
1. **"Error loading data from Google Sheets"**
   - Check if the Web App URL is correct in the HTML file
   - Ensure the Google Apps Script is deployed and accessible
   - Verify the Sheet ID in the Google Apps Script

2. **"Network error. Please try again."**
   - Check internet connection
   - Verify the Google Apps Script deployment is active

3. **"Invalid admin credentials" or "Invalid host credentials"**
   - Ensure you've set up the first host correctly
   - Check if the password and email are correct

### Testing the Setup
1. Try submitting player availability (name and email must match Players sheet)
2. Try updating attendance as host (name, email, password must match Hosts sheet)
3. Try adding a new player as admin
4. Try adding/removing hosts
5. Check if data appears in Google Sheets

## Maintenance
 - Regularly check Google Sheets for data integrity
 - Monitor the Google Apps Script execution transcript for errors
 - Update triggers if they stop working (re-run `setupTriggers()`)

## Support
If you encounter issues:
1. Check the Google Apps Script execution transcript for error details
2. Verify all permissions are granted
3. Ensure the Google Sheet is accessible and has the correct structure