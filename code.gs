/**
 * Sheet Ops Toolkit
 * -----------------
 * A small collection of Google Apps Script utilities that automate common
 * operations tasks inside Google Sheets: auditing calendar events, bulk
 * updating event descriptions, pulling file links from a Drive folder,
 * parsing JSON into rows, and resolving a human-readable name from a raw
 * URL (Drive, Forms, Colab, YouTube, Bitly).
 *
 * Built by Zakiyatun Surya (https://zakiyyah-ai.github.io) while managing
 * multi-program class operations, where these functions replaced hours of
 * manual cross-checking each week.
 *
 * Setup:
 * 1. Open (or create) a Google Sheet.
 * 2. Extensions > Apps Script, paste this file in as Code.gs, save.
 * 3. Reload the Sheet, a "Custom Function" menu will appear.
 * 4. (Optional, for calendar tools only) In the Apps Script editor:
 *    Services (+) > add "Google Calendar API" advanced service, and make
 *    sure the calendar you enter is one your account has access to.
 * 5. (Optional, for YouTube playlist name resolution only) Services (+) >
 *    add "YouTube Data API v3".
 *
 * See README.md for the expected sheet layout of each tool.
 *
 * License: MIT
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Function')
    .addItem('Get Links from Drive', 'GetDataDrive')
    .addItem('Parse Json', 'ParseJsonAndInsertData')
    .addItem('Get Name File from URL', 'getFileNameFromUrl')
    .addItem('List Filtered Events from Gcal', 'listFilteredEventsToSheet')
    .addItem('Bulk Update Event Descriptions', 'bulkUpdateEventDescriptions')
    .addToUi();
}

/**
 * Lists every file in a given Drive folder as a hyperlink + file name,
 * written into two columns you choose.
 */
function GetDataDrive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var c2 = Browser.inputBox('Enter cell for file names (e.g., A2)');
  var c1 = Browser.inputBox('Enter cell for URLs (e.g., B2)');
  var folderID = Browser.inputBox('Enter folder ID for import');

  if (c1 === '' || c2 === '' || folderID === '') {
    Browser.msgBox('One or more inputs are invalid');
    return;
  }

  var fldr;
  try {
    fldr = DriveApp.getFolderById(folderID);
  } catch (e) {
    Browser.msgBox('Folder not found or inaccessible: ' + folderID);
    return;
  }

  var c1Range = s.getRange(c1);
  var c2Range = s.getRange(c2);
  var files = fldr.getFiles();
  var urls = [];
  var names = [];
  var f, str;

  while (files.hasNext()) {
    f = files.next();
    str = '=hyperlink("' + f.getUrl() + '")';
    urls.push([str]);
    names.push([f.getName()]);
  }

  if (urls.length === 0) {
    Browser.msgBox('No files found in that folder.');
    return;
  }

  s.getRange(c1Range.getRow(), c1Range.getColumn(), urls.length).setFormulas(urls);
  s.getRange(c2Range.getRow(), c2Range.getColumn(), names.length).setValues(names);
}

/**
 * Reads a column of JSON strings (e.g. [{"label":"...","url":"..."}]) and
 * expands each into label/url pairs starting at a cell you choose. Each
 * item in the array gets 3 columns (label, url, blank spacer); rows with
 * more items than others just use more columns to the right — the output
 * width auto-expands to fit the widest row.
 */
function ParseJsonAndInsertData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  var dtrange = Browser.inputBox('Enter range of json (e.g., A:A)');
  if (dtrange === '' || dtrange === 'cancel') return;
  var dataRange = s.getRange(dtrange);
  var values = dataRange.getValues();

  var gtrange = Browser.inputBox('Enter starting cell for output json (e.g., T2:U)');
  if (gtrange === '' || gtrange === 'cancel') return;
  var startCell = s.getRange(gtrange);
  var startRow = startCell.getRow();
  var startColumn = startCell.getColumn();
  var numRows = dataRange.getNumRows();

  // Parse everything first so we know how many label/url pairs the widest
  // row needs, instead of assuming a fixed max of 2 items per row.
  var parsedRows = [];
  var maxItems = 1;
  for (var row = 0; row < values.length; row++) {
    var jsonString = values[row][0];
    if (!jsonString) {
      parsedRows.push(null);
      continue;
    }
    try {
      var data = JSON.parse(jsonString);
      parsedRows.push(data);
      if (data.length > maxItems) maxItems = data.length;
    } catch (e) {
      Logger.log('Error parsing JSON in row ' + (row + 1) + ': ' + e.message);
      parsedRows.push(null);
    }
  }

  var numColumns = maxItems * 3;
  s.getRange(startRow, startColumn, numRows, numColumns).clearContent();

  for (var r = 0; r < parsedRows.length; r++) {
    var rowData = parsedRows[r];
    if (!rowData) continue;
    for (var i = 0; i < rowData.length; i++) {
      var label = rowData[i].label;
      var url = rowData[i].url;
      var labelColumn = startColumn + i * 3;
      var linkColumn = startColumn + 1 + i * 3;
      s.getRange(startRow + r, labelColumn).setValue(label);
      s.getRange(startRow + r, linkColumn).setValue(url);
    }
  }
}

/**
 * For a column of URLs, tries to resolve a human-readable name: Drive file
 * name, Google Form title, Colab notebook name, Drive folder name, YouTube
 * playlist video titles, or the destination of a Bitly short link.
 */
function getFileNameFromUrl() {
  var linkColumn = Browser.inputBox('Enter column containing links (e.g., U)', 'Input Column', Browser.Buttons.OK_CANCEL);
  if (linkColumn === 'cancel' || linkColumn === '') {
    Browser.msgBox('No column selected. Operation canceled.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var range = s.getRange(linkColumn + '2:' + linkColumn + s.getLastRow());
  var values = range.getValues();

  var fileNames = [];
  for (var i = 0; i < values.length; i++) {
    fileNames.push([resolveNameFromLink(values[i][0])]);
  }

  var namesColumn = Browser.inputBox('Enter column for the name files (e.g., V)', 'Input Column', Browser.Buttons.OK_CANCEL);
  if (namesColumn === 'cancel' || namesColumn === '') {
    Browser.msgBox('No output column selected. Operation canceled.');
    return;
  }

  var outputRange = s.getRange(namesColumn + '2:' + namesColumn + (fileNames.length + 1));
  outputRange.setValues(fileNames);
}

function resolveNameFromLink(link) {
  if (!link) return '';
  if (link.includes('/bit.ly')) {
    return expandBitlyLink(link);
  }
  if (link.includes('/forms')) {
    return getGoogleFormTitle(link);
  }
  if (link.includes('https://colab.research.google.com/')) {
    return getGoogleDriveFileNameFromColabLink(link);
  }
  if (link.includes('https://drive.google.com/drive/folders/')) {
    var folderID = link.split('https://drive.google.com/drive/folders/')[1];
    return getGoogleFolderName(folderID);
  }
  if (link.includes('https://www.youtube.com/playlist?list=')) {
    var playlistId = link.split('https://www.youtube.com/playlist?list=')[1];
    return getYouTubeVideoTitles(playlistId);
  }
  if (link.includes('/d/')) {
    var fileId = link.split('/d/')[1].split('/')[0];
    return getGoogleDriveFileName(fileId);
  }
  return '';
}

function getGoogleFormTitle(formUrl) {
  try {
    var response = UrlFetchApp.fetch(formUrl);
    var html = response.getContentText();
    var titleMatch = html.match(/<title>([^<]*)<\/title>/);
    if (titleMatch && titleMatch[1]) {
      var title = titleMatch[1];
      return title.includes('Google Forms: Sign-in')
        ? 'Sign-in required to access the form'
        : title;
    }
    return 'Title not found';
  } catch (error) {
    return 'Form not found or inaccessible';
  }
}

function getGoogleDriveFileName(fileId) {
  try {
    return DriveApp.getFileById(fileId).getName();
  } catch (error) {
    return 'File not found or inaccessible';
  }
}

function getGoogleDriveFileIdFromColabLink(colabLink) {
  var fileIdMatch = colabLink.match(/drive\/(.*?)(\?|\/|$)/);
  return fileIdMatch && fileIdMatch[1] ? fileIdMatch[1] : null;
}

function getGoogleDriveFileNameFromColabLink(colabLink) {
  try {
    var fileId = getGoogleDriveFileIdFromColabLink(colabLink);
    return fileId ? DriveApp.getFileById(fileId).getName() : 'File ID not found in the link';
  } catch (error) {
    return 'File not found or inaccessible';
  }
}

function getGoogleFolderName(folderId) {
  try {
    return DriveApp.getFolderById(folderId).getName();
  } catch (error) {
    return 'Folder not found or inaccessible';
  }
}

function getYouTubeVideoTitles(playlistId) {
  try {
    var titles = [];
    var nextPageToken = '';
    var playlistTitle = '';
    var playlistResponse = YouTube.Playlists.list('snippet', { id: playlistId, part: 'snippet' });
    if (playlistResponse.items && playlistResponse.items.length > 0) {
      playlistTitle = playlistResponse.items[0].snippet.title;
    } else {
      return 'Error fetching playlist: Playlist not found.';
    }

    do {
      var response = YouTube.PlaylistItems.list('snippet', {
        playlistId: playlistId,
        part: 'snippet',
        maxResults: 50,
        pageToken: nextPageToken
      });
      for (var i = 0; i < response.items.length; i++) {
        titles.push(response.items[i].snippet.title);
      }
      nextPageToken = response.nextPageToken;
    } while (nextPageToken);

    return titles.length === 0
      ? 'The playlist "' + playlistTitle + '" is empty.'
      : titles.join(', ');
  } catch (error) {
    return 'Error fetching video titles: ' + error.message;
  }
}

function expandBitlyLink(bitlyURL) {
  try {
    var response = UrlFetchApp.fetch(bitlyURL, { followRedirects: false });
    return response.getHeaders()['Location'];
  } catch (error) {
    return 'Could not expand link';
  }
}

/**
 * Calendar Event Auditor: pulls every event in a date range from a chosen
 * calendar into a sheet, so you can spot-check invite lists, links, and
 * timing without opening Calendar event by event.
 *
 * Writes to a sheet you name, in the ACTIVE spreadsheet (no hardcoded IDs —
 * point this at any spreadsheet/calendar you have access to).
 */
function listFilteredEventsToSheet() {
  var calendarId = Browser.inputBox('Enter the calendar ID to audit (e.g., your email, or "primary")');
  if (calendarId === '' || calendarId === 'cancel') return;

  var sheetName = Browser.inputBox('Enter the sheet name to write results to (e.g., "gcal")');
  if (sheetName === '' || sheetName === 'cancel') return;

  var inputTimeMin = Browser.inputBox('Enter start date to filter (e.g., 2026-07-15T00:00:00Z)');
  var inputTimeMax = Browser.inputBox('Enter end date to filter (e.g., 2026-10-31T23:59:59Z)');
  var timeMin = new Date(inputTimeMin).toISOString();
  var timeMax = new Date(inputTimeMax).toISOString();

  var optionalArgs = {
    timeMin: timeMin,
    timeMax: timeMax,
    showDeleted: false,
    singleEvents: true,
    orderBy: 'startTime'
  };

  var events = Calendar.Events.list(calendarId, optionalArgs);
  if (!events.items || events.items.length === 0) {
    Browser.msgBox('No events found in that range.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow(['Event ID', 'Title', 'Description', 'Location/Link', 'Start Time', 'End Time', 'Invited Emails']);

  for (var i = 0; i < events.items.length; i++) {
    var event = events.items[i];
    var startTime = new Date(event.start.dateTime || event.start.date);
    var endTime = new Date(event.end.dateTime || event.end.date);
    var formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'dd MMM HH:mm:ss');
    var formattedEndTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'dd MMM HH:mm:ss');
    var attendees = event.attendees || [];
    var invitedEmails = attendees.map(function (a) { return a.email; }).join(', ');

    sheet.appendRow([
      event.id,
      event.summary,
      event.description,
      event.location,
      formattedStartTime,
      formattedEndTime,
      invitedEmails
    ]);
  }

  Browser.msgBox('Done — ' + events.items.length + ' events written to "' + sheetName + '".');
}

/**
 * Bulk Update Event Descriptions: reads a sheet of (title, link, start time,
 * description) rows and writes each description + link into the matching
 * Calendar event, matched by start time (to the minute) and exact title.
 *
 * NOTE: this REPLACES each event's existing description, it does not
 * append to it — export current descriptions first if you need to keep them.
 *
 * Sheet columns expected, in order, one row per event:
 *   A: Event title (must match exactly, case-insensitive)
 *   B: Link to append to the description
 *   C: Event start date/time (matched to the minute)
 *   D: New description text
 */
function bulkUpdateEventDescriptions() {
  var calendarId = Browser.inputBox('Enter the calendar ID to update (e.g., your email, or "primary")');
  if (calendarId === '' || calendarId === 'cancel') return;

  var sheetName = Browser.inputBox('Enter the sheet name with the update data (in this spreadsheet)');
  if (sheetName === '' || sheetName === 'cancel') return;

  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Browser.msgBox('Calendar not found or inaccessible: ' + calendarId);
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox('Sheet not found: ' + sheetName);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var updated = 0, skipped = 0;

  for (var i = 1; i < data.length; i++) {
    var titleSheet = String(data[i][0]).trim();
    var link = data[i][1];
    var startTime = new Date(data[i][2]);
    var description = data[i][3];

    if (!titleSheet || isNaN(startTime.getTime())) {
      Logger.log('Skipping row ' + (i + 1) + ': missing title or invalid start time.');
      skipped++;
      continue;
    }

    var events = calendar.getEvents(startTime, new Date(startTime.getTime() + 60000));
    var matches = events.filter(function (ev) {
      return ev.getStartTime().getTime() === startTime.getTime() &&
             ev.getTitle().trim().toLowerCase() === titleSheet.toLowerCase();
    });

    if (matches.length === 0) {
      Logger.log('❌ No match for "' + titleSheet + '" at ' + startTime);
      skipped++;
      continue;
    }
    if (matches.length > 1) {
      Logger.log('⚠️ ' + matches.length + ' matches for "' + titleSheet + '" at ' + startTime + ' — updating all.');
    }

    var newDesc = description + '\n\n🔗 Link: ' + link;
    matches.forEach(function (ev) {
      ev.setDescription(newDesc);
      updated++;
      Logger.log('✅ Updated: ' + titleSheet + ' | ' + startTime);
    });
  }

  Browser.msgBox('Done — updated ' + updated + ' event(s), skipped ' + skipped + ' row(s). Check Logs for details.');
}
