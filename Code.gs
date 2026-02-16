/**
 * Retrieves the timezone from the user's Google Calendar settings.
 * Returns a string like "Europe/Stockholm" or "America/New_York".
 */
function getUserTimezone() {
  try {
    return CalendarApp.getDefaultCalendar().getTimeZone();
  } catch (e) {
    // Fallback if permission is missing or API fails
    return "Etc/GMT";
  }
}

const DEFAULT_TIMEZONE = getUserTimezone();

function buildContextualCard(e) {
  var messageId = e.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var attachments = message.getAttachments();
  var icsFile = attachments.find(a => a.getName().toLowerCase().endsWith('.ics'));

  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle("Exchange Invite Preview")
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/event_black_24dp.png"));

  var section = CardService.newCardSection();

  if (!icsFile) {
    section.addWidget(CardService.newTextParagraph().setText("No .ics attachment found."));
    return card.addSection(section).build();
  }

  // 1. RAW DATA PREP
  var rawData = icsFile.getDataAsString();
  var unfoldedData = rawData.replace(/\r?\n[ \t]/g, "");

  // ISOLATE EVENT
  var eventBlock = unfoldedData.split("BEGIN:VEVENT")[1];
  if (!eventBlock) {
     section.addWidget(CardService.newTextParagraph().setText("No VEVENT block found."));
     return card.addSection(section).build();
  }

  // 2. PARSE FIELDS
  var summary = extractValue(eventBlock, "SUMMARY");
  var locationText = extractValue(eventBlock, "LOCATION");
  var description = extractValue(eventBlock, "DESCRIPTION");
  
  var dtStartLine = extractFullLine(eventBlock, "DTSTART");
  var dtEndLine = extractFullLine(eventBlock, "DTEND");

  // 3. PARSE MEETING URLS
  var teamsUrl = extractValue(eventBlock, "X-MICROSOFT-SKYPETEAMSMEETINGURL");
  var onlineUrl = extractValue(eventBlock, "X-MICROSOFT-ONLINEMEETINGEXTERNALLINK");
  var meetingLink = (teamsUrl || onlineUrl || "").trim();

  // 4. SMART LOCATION LOGIC
  var finalLocation = locationText;
  if (meetingLink) {
    if (!locationText || locationText.toLowerCase().includes("teams meeting")) {
      finalLocation = meetingLink;
    } else {
      finalLocation = locationText + " (" + meetingLink + ")";
    }
  }

  // Clean up description (newlines and commas)
  description = description.replace(/\\n/g, "\n").replace(/\\,/g, ",");

  // 5. PARSE DATES (Timezone Aware)
  var startDate = parseIcsLineToDate(dtStartLine);
  var endDate = parseIcsLineToDate(dtEndLine);

  if (!startDate) {
     section.addWidget(CardService.newTextParagraph().setText("Could not parse start date."));
     return card.addSection(section).build();
  }
  
  if (!endDate) {
    endDate = new Date(startDate.getTime() + (60 * 60 * 1000)); 
  }

  // 6. BUILD UI
  var displayString = Utilities.formatDate(startDate, DEFAULT_TIMEZONE, "yyyy-MM-dd HH:mm") + ' - ' + Utilities.formatDate(endDate, DEFAULT_TIMEZONE, "HH:mm")

  // A. Event Title
  section.addWidget(CardService.newDecoratedText()
    .setTopLabel("Event")
    .setText(summary || "Untitled")
    .setWrapText(true));

  // B. Time
  section.addWidget(CardService.newDecoratedText()
    .setTopLabel("Start Time (" + DEFAULT_TIMEZONE + ")")
    .setText(displayString)
    .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK)));
    
  // C. Online Meeting Link (If exists)
  if (meetingLink) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel("Online Meeting")
      .setText("Join Link Detected")
      .setOpenLink(CardService.newOpenLink().setUrl(meetingLink))
      .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.VIDEO_CAMERA)));
  } 

  // D. Physical Location (If exists AND isn't just "Teams Meeting")
  if (locationText && !locationText.toLowerCase().includes("teams meeting")) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel("Location")
      .setText(locationText)
      .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.MAP_PIN)));
  }

  // E. Description (ALWAYS show if present)
  if (description) {
    // Truncate if extremely long to save UI space (optional)
    var descPreview = description.length > 300 ? description.substring(0, 297) + "..." : description;
    
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel("Description")
      .setText(descPreview)
      .setWrapText(true)
      .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION)));
  }

  // F. Add Button
  var action = CardService.newAction()
    .setFunctionName('createCalendarEvent')
    .setParameters({
      summary: summary, 
      location: finalLocation, 
      description: description, // Passes full description to calendar
      startTimeMs: startDate.getTime().toString(),
      endTimeMs: endDate.getTime().toString()
    });

  section.addWidget(CardService.newTextButton()
    .setText("Add to Calendar")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(action));

  return card.addSection(section).build();
}

// --- HELPER FUNCTIONS ---

function extractValue(data, tag) {
  var regex = new RegExp("^" + tag + "(?:;.*?)?:(.*)$", "m");
  var match = data.match(regex);
  return match ? match[1].trim() : "";
}

function extractFullLine(data, tag) {
  var regex = new RegExp("^" + tag + ".*$", "m");
  var match = data.match(regex);
  return match ? match[0].trim() : "";
}

function parseIcsLineToDate(line) {
  if (!line) return null;
  var valueParts = line.split(":");
  var rawDate = valueParts.pop(); 
  
  var tzid = "Etc/GMT"; // Default fallback
  var tzidMatch = line.match(/TZID=([^:;]+)/);
  if (tzidMatch) {
    tzid = resolveTzid(tzidMatch[1]);
  } else if (rawDate.endsWith("Z")) {
    tzid = "Etc/GMT";
  }

  var digits = rawDate.match(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})/);
  if (!digits) return null;

  var year = digits[1];
  var month = digits[2];
  var day = digits[3];
  var hour = digits[4];
  var min = digits[5];
  var sec = digits[6];

  var dateString = year + "-" + month + "-" + day + " " + hour + ":" + min + ":" + sec;
  
  // Use the timezone from the file to calculate absolute time
  return Utilities.parseDate(dateString, tzid, "yyyy-MM-dd HH:mm:ss");
}

function resolveTzid(msTzid) {
  var cleanTz = msTzid.replace(/"/g, "").trim();
  var map = {
    "GMT Standard Time": "Europe/London",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central Europe Standard Time": "Europe/Warsaw",
    "Romance Standard Time": "Europe/Paris",
    "GTB Standard Time": "Europe/Bucharest",
    "E. Europe Standard Time": "Europe/Chisinau",
    "Egypt Standard Time": "Africa/Cairo",
    "Fle Standard Time": "Europe/Kiev",
    "Israel Standard Time": "Asia/Jerusalem",
    "South Africa Standard Time": "Africa/Johannesburg",
    "Russian Standard Time": "Europe/Moscow",
    "Arab Standard Time": "Asia/Riyadh",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Pacific Standard Time": "America/Los_Angeles",
    "India Standard Time": "Asia/Kolkata",
    "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo",
    "AUS Eastern Standard Time": "Australia/Sydney"
  };
  return map[cleanTz] || cleanTz;
}

function createCalendarEvent(e) {
  var p = e.parameters;
  var start = new Date(parseInt(p.startTimeMs));
  var end = new Date(parseInt(p.endTimeMs));

  CalendarApp.getDefaultCalendar().createEvent(p.summary, start, end, {
    location: p.location,
    description: p.description
  });
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Event added successfully!"))
    .build();
}
