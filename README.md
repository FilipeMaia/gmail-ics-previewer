# Gmail ICS Calendar Add-on (Exchange Fix)

A vibe-coded private Google Workspace Add-on that forces Gmail to handle `.ics` calendar invitations correctly, even when they are synced via IMAP or forwarded from Microsoft Exchange.

##  The Problem
If you sync your work email (Exchange/Outlook) to Gmail using **IMAP**, Google Calendar often treats invitations as static file attachments rather than interactive events.
*  No "Yes/No/Maybe" buttons.
*  No "Add to Calendar" button.
*  Timezone issues (events showing at wrong times).
*  Missing Microsoft Teams/Zoom links.

##  The Solution
This Google Apps Script creates a **custom sidebar** in Gmail. When you open an email with an `.ics` attachment:
1.  It reads the file and **previews** the event details (Title, Time, Location).
2.  It intelligently **parses Microsoft Teams/Zoom links** and puts them in the Location field.
3.  It handles **Timezone Mapping** (fixing the "GMT Standard Time" vs. "Europe/London" issue).
4.  It allows you to **Add to Calendar** with a single click.

##  Features
* **Smart Link Detection:** Automatically extracts `X-MICROSOFT-SKYPETEAMSMEETINGURL` and puts it in the location field so it's clickable.
* **Timezone Aware:** correctly maps Windows timezones (e.g., "Romance Standard Time") to IANA timezones (e.g., "Europe/Paris") to prevent events from being 1 hour off.
* **Robust Parsing:** Handles line-folding (split lines) in `.ics` files that often break other parsers.
* **Privacy First:** Runs entirely within your own Google Account. No data is sent to third-party servers.

---

##  Installation Guide

You do not need to install anything on your computer. This script runs in the Google Cloud.

### Step 1: Create the Project
1.  Go to [script.google.com](https://script.google.com).
2.  Click **New Project**.
3.  Name the project: `Gmail ICS Add-on`.

### Step 2: Update the Manifest (`appsscript.json`)
By default, the manifest file is hidden.
1.  Click the **Project Settings** ( gear icon) on the left sidebar.
2.  Check the box **"Show 'appsscript.json' manifest file in editor"**.
3.  Go back to the **Editor** (code `< >` icon).
4.  Click on `appsscript.json` in the file list.
5.  **Replace** the entire content with the code below:

```json
{
  "timeZone": "Europe/Stockholm",
  "dependencies": {
    "enabledAdvancedServices": []
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "[https://www.googleapis.com/auth/gmail.addons.execute](https://www.googleapis.com/auth/gmail.addons.execute)",
    "[https://www.googleapis.com/auth/gmail.readonly](https://www.googleapis.com/auth/gmail.readonly)",
    "[https://www.googleapis.com/auth/calendar](https://www.googleapis.com/auth/calendar)"
  ],
  "gmail": {
    "name": "ICS Previewer",
    "logoUrl": "[https://www.gstatic.com/images/icons/material/system/1x/event_black_24dp.png](https://www.gstatic.com/images/icons/material/system/1x/event_black_24dp.png)",
    "contextualTriggers": [{
      "unconditional": {},
      "onTriggerFunction": "buildContextualCard"
    }],
    "primaryColor": "#4285F4"
  }
}
