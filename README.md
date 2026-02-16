# Gmail ICS Calendar Add-on (Exchange Fix)

A private Google Workspace Add-on that forces Gmail to handle `.ics` calendar invitations correctly, even when they are synced via IMAP or forwarded from Microsoft Exchange.

##  The Problem
If you sync your work email (Exchange/Outlook) to Gmail using **IMAP**, Google Calendar often treats invitations as static file attachments rather than interactive events.
*  No "Add to Calendar" button.
*  Timezone issues (events showing at wrong times).
*  Missing Teams/Zoom links.

##  The Solution
This Google Apps Script creates a **custom sidebar** in Gmail. When you open an email with an `.ics` attachment:
1.  It reads the file and **previews** the event details (Title, Time, Location).
2.  It allows you to **Add to Calendar** with a single click.
3.  It intelligently **parses Teams/Zoom links** and puts them in the Location field.
4.  It handles **Timezone Mapping**.
5. **Privacy First:** Runs entirely within your own Google Account. No data is sent to third-party servers.


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

```

### Step 3: Add the Code (`Code.gs`)

1. Click on `Code.gs` in the file list.
2. **Delete** any placeholder code (like `function myFunction...`).
3. **Copy and Paste** the full script from `Code.gs` in this repository.
* *Note: Ensure you set the `const DEFAULT_TIMEZONE` at the top of the file to your local timezone.*



### Step 4: Deploy & Authorize

1. Click the blue **Deploy** button (top right) > **Test deployments**.
2. Click **Install** (next to the "Gmail" application type).
3. Click **Done**.
4. Open **Gmail** in a new tab (or refresh if open).
5. Open any email that has an `.ics` attachment.
6. Look for the **Event Icon** () in the right-hand sidebar.
7. Click it. You will be asked to **Authorize** the script. Click "Review Permissions" and allow it.


##  Troubleshooting

**"Specified permissions are not sufficient"**

* You likely didn't update `appsscript.json` correctly. Make sure `.../auth/calendar` is listed in the `oauthScopes` section, not just `calendar.events`.

**The sidebar doesn't appear**

* Ensure you are opening an email that actually has a file ending in `.ics` attached. The sidebar only triggers when that file type is detected.

**The date shows as 1601**

* This means the parser grabbed the "Timezone Definition" date instead of the Event date. Ensure you are using the latest version of `Code.gs` from this repo, which includes the `split("BEGIN:VEVENT")` fix.

##  License

See the LICENSE file.

