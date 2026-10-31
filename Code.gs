/**
 * HR CRM SYSTEM
 * Google Apps Script Backend
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const HR_SHEET = 'HR_Data';
const USER_SHEET = 'Users';
const REMARKS_SHEET = 'Remarks_Log';
const ACTIVITY_SHEET = 'Activity_Log';
const EVENTS_SHEET = 'Events';

// Public Google calendar of Indian government holidays — always current,
// no manual list to maintain.
const HOLIDAY_CALENDAR_ID = 'en.indian#holiday@group.v.calendar.google.com';

// Script Properties key storing the JSON array of emails subscribed to
// the daily "2 days before" event-reminder email.
const REMINDER_SUBS_KEY = 'REMINDER_SUBSCRIBERS';

/* =========================
   INITIAL SETUP
========================= */

function setupSheets() {

  // HR DATA SHEET
  let hrSheet = SS.getSheetByName(HR_SHEET);

  if (!hrSheet) {

    hrSheet = SS.insertSheet(HR_SHEET);

    hrSheet.appendRow([
      'ID',
      'HR Name',
      'Company Name',
      'Phone Number',
      'Email ID',
      'Last Contact',
      'Next Follow-up',
      'Status',
      'Remarks',
      'Designation',
      'First Call',
      'Call Script'
    ]);

  } else {

    // MIGRATION: add any columns that don't exist yet on older sheets
    // (keeps existing columns / data untouched)
    ensureColumn(hrSheet, 'Designation');
    ensureColumn(hrSheet, 'First Call');
    ensureColumn(hrSheet, 'Call Script');
  }

  // MIGRATION: backfill a stable UUID for any row missing one in column A.
  // This is critical — every lookup (remarks, edits, deletes) now keys off
  // this ID instead of the row number, since row numbers shift after deletes.
  const lastRow = hrSheet.getLastRow();

  if (lastRow > 1) {

    const idRange = hrSheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idRange.getValues();
    let changed = false;

    for (let i = 0; i < ids.length; i++) {

      if (!ids[i][0]) {
        ids[i][0] = Utilities.getUuid();
        changed = true;
      }
    }

    if (changed) {
      idRange.setValues(ids);
    }
  }

  // USER SHEET
  let userSheet = SS.getSheetByName(USER_SHEET);

  if (!userSheet) {

    userSheet = SS.insertSheet(USER_SHEET);

    userSheet.appendRow([
      'Email',
      'Password'
    ]);

    userSheet.appendRow([
      'admin@example.com',
      'admin123'
    ]);
  }

  // REMARKS LOG SHEET
  let remarksSheet = SS.getSheetByName(REMARKS_SHEET);

  if (!remarksSheet) {

    remarksSheet = SS.insertSheet(REMARKS_SHEET);

    remarksSheet.appendRow([
      'ID',
      'RowID',
      'Timestamp',
      'Remark'
    ]);
  }

  // ACTIVITY LOG SHEET (tracks who did what, for the in-app activity feed)
  let activitySheet = SS.getSheetByName(ACTIVITY_SHEET);

  if (!activitySheet) {

    activitySheet = SS.insertSheet(ACTIVITY_SHEET);

    activitySheet.appendRow([
      'ID',
      'Timestamp',
      'User',
      'Action',
      'Details'
    ]);
  }

  // EVENTS SHEET (calendar: campus drives / other events)
  let eventsSheet = SS.getSheetByName(EVENTS_SHEET);

  if (!eventsSheet) {

    eventsSheet = SS.insertSheet(EVENTS_SHEET);

    eventsSheet.appendRow([
      'ID',
      'Event Name',
      'Company',
      'Type',
      'Branches',
      'Event Date',
      'Status',
      'Notes',
      'Created By',
      'Created At'
    ]);
  }
}

/* =========================
   HELPER: add a header column if it doesn't already exist
   (keeps older spreadsheets migrating forward safely)
========================= */

function ensureColumn(sheet, headerName) {

  const lastCol = sheet.getLastColumn();

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers.indexOf(headerName) === -1) {
    sheet.getRange(1, lastCol + 1).setValue(headerName);
  }
}

/* =========================
   HELPER: resolve current row number from a stable UUID
   (row numbers shift after deletes, so we never trust them directly)
========================= */

function getRowNumberByUuid(sheet, uuid) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {

    if (String(ids[i][0]) === String(uuid)) {
      return i + 2;
    }
  }

  return -1;
}

/* =========================
   WEB APP
========================= */

function doGet() {

  setupSheets();

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'CRM API' }))
    .setMimeType(ContentService.MimeType.JSON);
}



/* =========================
   JSON API (used by the Vercel front end)
========================= */

const API_FUNCTIONS = {
  authenticate, getHRData, quickUpdate, postponeAllOverdue, updateHRRemark,
  addRemarkLog, getRemarksForRow, editRemarkLog, updateHRDetails,
  getCommonScript, updateCommonScript, toggleFirstCall, updateCallScript,
  addHR, deleteHR, getRecentActivity,
  getEvents, addEvent, updateEvent, cancelEvent, restoreEvent, deleteEvent,
  getUpcomingReminders, getHolidays,
  isSubscribedToReminders, subscribeToReminders, unsubscribeFromReminders
};

const TOKEN_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

function getApiSecret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('API_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('API_SECRET', s);
  }
  return s;
}

function sign_(payload) {
  const sig = Utilities.computeHmacSha256Signature(payload, getApiSecret_());
  return Utilities.base64EncodeWebSafe(sig);
}

function makeToken_(email) {
  const payload = email + '|' + (Date.now() + TOKEN_TTL_MS);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sign_(payload);
}

function verifyToken_(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length !== 2) return false;
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    if (sign_(payload) !== parts[1]) return false;
    return Number(payload.split('|').pop()) > Date.now();
  } catch (err) {
    return false;
  }
}

function doPost(e) {

  let out;

  try {

    const req = JSON.parse(e.postData.contents);
    const fn = req.fn;
    const args = req.args || [];

    if (!Object.prototype.hasOwnProperty.call(API_FUNCTIONS, fn)) {
      throw new Error('Unknown function: ' + fn);
    }

    if (fn !== 'authenticate' && !verifyToken_(req.token)) {
      out = { error: 'unauthorized' };
    } else {

      let result = API_FUNCTIONS[fn].apply(null, args);

      if (fn === 'authenticate' && result && result.success) {
        result = { success: true, email: result.email, token: makeToken_(result.email) };
      }

      out = { result: result === undefined ? null : result };
    }

  } catch (err) {
    out = { error: String(err && err.message || err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   LOGIN
========================= */

function authenticate(email, password) {

  const sheet = SS.getSheetByName(USER_SHEET);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const dbEmail = String(data[i][0]).trim();
    const dbPassword = String(data[i][1]).trim();

    if (dbEmail === email && dbPassword === password) {

      return {
        success: true,
        email: email
      };
    }
  }

  return {
    success: false
  };
}

/* =========================
   GET HR DATA
========================= */

function getHRData() {

  const sheet = SS.getSheetByName(HR_SHEET);

  const data = sheet.getDataRange().getValues();

  data.shift();

  const today = new Date().setHours(0,0,0,0);

  return data.map((row, index) => {

    const nextDate = row[6] ? new Date(row[6]) : null;

    const status = row[7] || 'Pending';

    // A contact is only "due" while its follow-up date IS today.
    // Once that day passes with no new remark/date, it naturally stops
    // matching here and drops out of Priority — no leftover reminders.
    const isOverdue =
      nextDate &&
      nextDate.setHours(0,0,0,0) === today &&
      status !== 'Not Interested';

    const isPriority = isOverdue;

    return {

      rowId: row[0],

      hrName: row[1] || '',

      company: row[2] || '',

      phone: row[3] || '',

      email: row[4] || '',

      lastContact: row[5]
        ? Utilities.formatDate(
            new Date(row[5]),
            Session.getScriptTimeZone(),
            'dd/MM/yyyy'
          )
        : '',

      nextFollowUp: row[6]
        ? Utilities.formatDate(
            new Date(row[6]),
            Session.getScriptTimeZone(),
            'dd/MM/yyyy'
          )
        : '',

      // ISO format (yyyy-MM-dd) kept separately so the <input type="date">
      // field in the details modal can be populated correctly
      nextFollowUpISO: row[6]
        ? Utilities.formatDate(
            new Date(row[6]),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          )
        : '',

      status: status,

      remarks: row[8] || '',

      designation: row[9] || '',

      firstCall: row[10] === true || row[10] === 'Yes',

      callScript: row[11] || '',

      isOverdue: isOverdue,

      isPriority: isPriority
    };
  });
}

/* =========================
   QUICK STATUS UPDATE
========================= */

function quickUpdate(uuid, status) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  const today = new Date();

  // Last Contact
  sheet.getRange(rowId, 6).setValue(today);

  // Status
  sheet.getRange(rowId, 8).setValue(status);

  // Auto callback follow-up
  if (status === 'Call Back') {

    const nextDate = new Date();

    nextDate.setDate(today.getDate() + 2);

    sheet.getRange(rowId, 7).setValue(nextDate);
  }

  return {
    success: true
  };
}

/* =========================
   POSTPONE ALL PRIORITY / OVERDUE (bulk +1 day)
========================= */

function postponeAllOverdue(userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { success: true, count: 0 };
  }

  // Columns: A ID .. G Next Follow-up, H Status
  const range = sheet.getRange(2, 1, lastRow - 1, 8);
  const data = range.getValues();

  const today = new Date().setHours(0, 0, 0, 0);

  let count = 0;

  for (let i = 0; i < data.length; i++) {

    const row = data[i];
    const nextDateRaw = row[6];
    const status = row[7] || 'Pending';

    if (!nextDateRaw) continue;

    // Same "is this due today" check used by getHRData / isOverdue
    const checkDate = new Date(nextDateRaw);
    const isOverdue =
      checkDate.setHours(0, 0, 0, 0) === today &&
      status !== 'Not Interested';

    if (isOverdue) {

      const newDate = new Date(nextDateRaw);
      newDate.setDate(newDate.getDate() + 1);

      sheet.getRange(i + 2, 7).setValue(newDate);

      count++;
    }
  }

  if (count > 0) {
    logActivity(
      userEmail,
      'Postponed Follow-ups',
      `Postponed ${count} priority/overdue contact${count > 1 ? 's' : ''} by 1 day`
    );
  }

  return {
    success: true,
    count: count
  };
}

/* =========================
   UPDATE REMARK
========================= */

function updateHRRemark(uuid, newRemark, nextCallDate, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  const contactInfo = sheet.getRange(rowId, 2, 1, 2).getValues()[0];
  const contactLabel = `${contactInfo[0] || 'Unnamed'} at ${contactInfo[1] || 'Unknown company'}`;

  // Save latest remark for quick reference, and log full history
  if (newRemark && newRemark.trim() !== '') {

    sheet.getRange(rowId, 9).setValue(newRemark.trim());

    addRemarkLog(uuid, newRemark.trim());

    logActivity(userEmail, 'New Remark', contactLabel);
  }

  // Save next call date
  if (nextCallDate) {

    sheet.getRange(rowId, 7).setValue(
      new Date(nextCallDate)
    );

    logActivity(userEmail, 'Follow-up Scheduled', contactLabel);
  }

  // Update last contact
  sheet.getRange(rowId, 6).setValue(
    new Date()
  );

  return {
    success: true
  };
}

/* =========================
   REMARKS HISTORY
========================= */

// Add a new entry to the remarks log
function addRemarkLog(uuid, text) {

  const sheet = SS.getSheetByName(REMARKS_SHEET);

  sheet.appendRow([
    Utilities.getUuid(),
    uuid,
    new Date(),
    text
  ]);
}

// Get remarks for a given HR contact (by stable UUID), most recent first
function getRemarksForRow(uuid) {

  const sheet = SS.getSheetByName(REMARKS_SHEET);

  const data = sheet.getDataRange().getValues();

  data.shift();

  const rows = data.filter(r => String(r[1]) === String(uuid));

  rows.sort((a, b) => new Date(b[2]) - new Date(a[2]));

  return rows.map(r => ({
    id: r[0],
    timestamp: Utilities.formatDate(
      new Date(r[2]),
      Session.getScriptTimeZone(),
      'dd MMM yyyy, hh:mm a'
    ),
    text: r[3]
  }));
}

// Edit the text of an existing remark
function editRemarkLog(remarkId, newText, userEmail) {

  const sheet = SS.getSheetByName(REMARKS_SHEET);

  const data = sheet.getDataRange().getValues();

  let targetUuid = null;

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]) === String(remarkId)) {

      sheet.getRange(i + 1, 4).setValue(newText);

      targetUuid = data[i][1];

      break;
    }
  }

  // Keep the HR_Data "quick view" remark column in sync
  // if the edited remark is the most recent one for that contact
  if (targetUuid) {

    const remarks = getRemarksForRow(targetUuid);

    if (remarks.length > 0) {

      const hrSheet = SS.getSheetByName(HR_SHEET);

      const rowId = getRowNumberByUuid(hrSheet, targetUuid);

      if (rowId !== -1) {

        hrSheet.getRange(rowId, 9).setValue(remarks[0].text);

        const contactInfo = hrSheet.getRange(rowId, 2, 1, 2).getValues()[0];

        logActivity(
          userEmail,
          'Edited Remark',
          `${contactInfo[0] || 'Unnamed'} at ${contactInfo[1] || 'Unknown company'}`
        );
      }
    }
  }

  return {
    success: true
  };
}

/* =========================
   UPDATE HR DETAILS (EDIT)
========================= */

function updateHRDetails(uuid, obj, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  sheet.getRange(rowId, 2).setValue(obj.hrName || '');

  sheet.getRange(rowId, 3).setValue(obj.company || '');

  sheet.getRange(rowId, 4).setValue(obj.phone || '');

  sheet.getRange(rowId, 5).setValue(obj.email || '');

  sheet.getRange(rowId, 10).setValue(obj.designation || '');

  logActivity(
    userEmail,
    'Edited Details',
    `${obj.hrName || 'Unnamed'} at ${obj.company || 'Unknown company'}`
  );

  return {
    success: true
  };
}

/* =========================
   COMMON CALL SCRIPT
   One shared script for every lead, opened from a single button in the
   UI (top-right, next to "+ Add Lead") instead of a per-card button.
   Stored once in Script Properties rather than per-row.
========================= */

const COMMON_SCRIPT_KEY = 'COMMON_CALL_SCRIPT';

// Return the shared call script (HTML), or '' if none has been saved yet.
function getCommonScript() {

  return PropertiesService.getScriptProperties().getProperty(COMMON_SCRIPT_KEY) || '';
}

// Save the shared call script (HTML) used by everyone.
function updateCommonScript(scriptHtml, userEmail) {

  PropertiesService.getScriptProperties().setProperty(COMMON_SCRIPT_KEY, scriptHtml || '');

  logActivity(userEmail, 'Updated Common Script', 'Shared call script updated');

  return { success: true };
}

/* =========================
   1ST CALL TOGGLE + CALL SCRIPT (legacy, per-lead — no longer surfaced
   in the UI, kept only so existing sheet data/columns aren't disturbed)
========================= */

// Flip the "First Call" flag for a contact (shows/hides the script box on the card)
function toggleFirstCall(uuid, isFirstCall, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  sheet.getRange(rowId, 11).setValue(isFirstCall ? true : false);

  const contactInfo = sheet.getRange(rowId, 2, 1, 2).getValues()[0];
  const contactLabel = `${contactInfo[0] || 'Unnamed'} at ${contactInfo[1] || 'Unknown company'}`;

  logActivity(
    userEmail,
    isFirstCall ? 'Marked 1st Call' : 'Unmarked 1st Call',
    contactLabel
  );

  return { success: true };
}

// Save the editable/formattable talking-points script for a contact (stored as HTML)
function updateCallScript(uuid, script, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  sheet.getRange(rowId, 12).setValue(script || '');

  const contactInfo = sheet.getRange(rowId, 2, 1, 2).getValues()[0];

  logActivity(
    userEmail,
    'Updated Call Script',
    `${contactInfo[0] || 'Unnamed'} at ${contactInfo[1] || 'Unknown company'}`
  );

  return { success: true };
}

/* =========================
   ADD NEW HR
========================= */

function addHR(obj, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const newId = Utilities.getUuid();

  sheet.appendRow([

    newId,

    obj.hrName || '',

    obj.company || '',

    obj.phone || '',

    obj.email || '',

    '',

    obj.nextFollowUp || '',

    'Pending',

    obj.remarks || '',

    obj.designation || '',

    obj.firstCall ? true : false,

    obj.firstCall ? (obj.callScript || '') : ''
  ]);

  if (obj.remarks && obj.remarks.trim() !== '') {

    addRemarkLog(newId, obj.remarks.trim());
  }

  logActivity(
    userEmail,
    'New Lead',
    `${obj.hrName || 'Unnamed'} at ${obj.company || 'Unknown company'}`
  );

  return {
    success: true
  };
}

/* =========================
   DELETE HR
========================= */

function deleteHR(uuid, userEmail) {

  const sheet = SS.getSheetByName(HR_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  // Capture contact info before deleting the row, so we can log it
  const contactInfo = sheet.getRange(rowId, 2, 1, 2).getValues()[0];
  const contactLabel = `${contactInfo[0] || 'Unnamed'} at ${contactInfo[1] || 'Unknown company'}`;

  sheet.deleteRow(rowId);

  // Clean up: remove any remarks that were logged against this contact,
  // so no orphaned history is left behind in Remarks_Log
  deleteRemarksForUuid(uuid);

  logActivity(userEmail, 'Deleted Contact', contactLabel);

  return {
    success: true
  };
}

// Delete every remarks-log entry tied to a given contact UUID
function deleteRemarksForUuid(uuid) {

  const sheet = SS.getSheetByName(REMARKS_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  // Delete from the bottom up so row indices don't shift mid-loop
  for (let i = data.length - 1; i >= 0; i--) {

    if (String(data[i][1]) === String(uuid)) {

      sheet.deleteRow(i + 2);
    }
  }
}

/* =========================
   ACTIVITY FEED
   (shows what teammates have changed — added leads, edits, remarks, deletes)
========================= */

// Record one activity entry. userEmail may be blank/unknown if not supplied.
function logActivity(userEmail, action, details) {

  const sheet = SS.getSheetByName(ACTIVITY_SHEET);

  if (!sheet) return;

  sheet.appendRow([
    Utilities.getUuid(),
    new Date(),
    userEmail || 'Unknown',
    action,
    details || ''
  ]);
}

// Return the most recent activity entries, newest first.
// Used by the frontend to populate the "Recent Activity" panel and badge count.
function getRecentActivity(limit) {

  const sheet = SS.getSheetByName(ACTIVITY_SHEET);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const max = limit || 50;

  return data
    .map(r => ({
      id: r[0],
      timestampRaw: new Date(r[1]).getTime(),
      timestamp: Utilities.formatDate(
        new Date(r[1]),
        Session.getScriptTimeZone(),
        'dd MMM yyyy, hh:mm a'
      ),
      user: r[2],
      action: r[3],
      details: r[4]
    }))
    .sort((a, b) => b.timestampRaw - a.timestampRaw)
    .slice(0, max);
}

/* =========================
   EVENTS CALENDAR
   Campus drives / other events: add, reschedule, cancel, delete.
   Government holidays are pulled live from Google's public Indian
   holiday calendar (no manual list to keep updated).
========================= */

// Return every event, newest-dated first not required — frontend sorts/filters.
function getEvents() {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const data = sheet.getDataRange().getValues();

  data.shift();

  return data.map(row => ({

    id: row[0],

    eventName: row[1] || '',

    company: row[2] || '',

    type: row[3] || 'Other',

    branches: row[4] || '',

    eventDateISO: row[5]
      ? Utilities.formatDate(
          new Date(row[5]),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )
      : '',

    status: row[6] || 'Scheduled',

    notes: row[7] || '',

    createdBy: row[8] || '',

    createdAt: row[9]
      ? Utilities.formatDate(
          new Date(row[9]),
          Session.getScriptTimeZone(),
          'dd MMM yyyy, hh:mm a'
        )
      : ''
  }));
}

// Add a new event ("mark" an event on the calendar).
// obj: { eventName, company, isCampusDrive, branches: [...], eventDate: 'yyyy-MM-dd', notes }
function addEvent(obj, userEmail) {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const newId = Utilities.getUuid();

  const type = obj.isCampusDrive ? 'Campus Drive' : 'Other';
  const branches = obj.isCampusDrive ? (obj.branches || []).join(', ') : '';

  sheet.appendRow([
    newId,
    obj.eventName || '',
    obj.company || '',
    type,
    branches,
    obj.eventDate ? new Date(obj.eventDate) : '',
    'Scheduled',
    obj.notes || '',
    userEmail || '',
    new Date()
  ]);

  logActivity(
    userEmail,
    'New Event',
    `${obj.eventName || 'Untitled'}${obj.eventDate ? ' on ' + obj.eventDate : ''}`
  );

  return { success: true, id: newId };
}

// Edit an event's details — also how a reschedule (date change) is saved.
function updateEvent(uuid, obj, userEmail) {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  const oldDateRaw = sheet.getRange(rowId, 6).getValue();

  const oldDateISO = oldDateRaw
    ? Utilities.formatDate(new Date(oldDateRaw), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';

  const type = obj.isCampusDrive ? 'Campus Drive' : 'Other';
  const branches = obj.isCampusDrive ? (obj.branches || []).join(', ') : '';

  sheet.getRange(rowId, 2).setValue(obj.eventName || '');
  sheet.getRange(rowId, 3).setValue(obj.company || '');
  sheet.getRange(rowId, 4).setValue(type);
  sheet.getRange(rowId, 5).setValue(branches);
  sheet.getRange(rowId, 6).setValue(obj.eventDate ? new Date(obj.eventDate) : '');
  sheet.getRange(rowId, 8).setValue(obj.notes || '');

  const rescheduled = !!obj.eventDate && obj.eventDate !== oldDateISO;

  logActivity(
    userEmail,
    rescheduled ? 'Rescheduled Event' : 'Updated Event',
    `${obj.eventName || 'Untitled'}${rescheduled ? ' → ' + obj.eventDate : ''}`
  );

  return { success: true };
}

// Cancel an event (kept in the sheet, shown struck-through, no reminders sent for it).
function cancelEvent(uuid, userEmail) {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  sheet.getRange(rowId, 7).setValue('Cancelled');

  const name = sheet.getRange(rowId, 2).getValue();

  logActivity(userEmail, 'Cancelled Event', name || 'Untitled');

  return { success: true };
}

// Undo a cancellation.
function restoreEvent(uuid, userEmail) {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  sheet.getRange(rowId, 7).setValue('Scheduled');

  const name = sheet.getRange(rowId, 2).getValue();

  logActivity(userEmail, 'Restored Event', name || 'Untitled');

  return { success: true };
}

// Permanently remove an event.
function deleteEvent(uuid, userEmail) {

  const sheet = SS.getSheetByName(EVENTS_SHEET);

  const rowId = getRowNumberByUuid(sheet, uuid);

  if (rowId === -1) {
    return { success: false };
  }

  const name = sheet.getRange(rowId, 2).getValue();

  sheet.deleteRow(rowId);

  logActivity(userEmail, 'Deleted Event', name || 'Untitled');

  return { success: true };
}

// Events happening today, tomorrow, or the day after (still Scheduled) —
// powers the in-app reminder banner and the calendar bell badge.
function getUpcomingReminders() {

  const events = getEvents();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const limit = new Date(today);
  limit.setDate(limit.getDate() + 2);

  return events
    .filter(e => {

      if (e.status !== 'Scheduled' || !e.eventDateISO) return false;

      const d = new Date(e.eventDateISO + 'T00:00:00');

      return d >= today && d <= limit;
    })
    .sort((a, b) => a.eventDateISO.localeCompare(b.eventDateISO));
}

/* =========================
   GOVERNMENT HOLIDAYS
   Pulled live from Google's public Indian holiday calendar and cached
   for a few hours so repeated calendar views stay fast.
========================= */

function getHolidays(year) {

  const cache = CacheService.getScriptCache();
  const cacheKey = 'holidays_' + year;

  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  let holidays = [];

  try {

    const cal = CalendarApp.getCalendarById(HOLIDAY_CALENDAR_ID);

    if (cal) {

      const start = new Date(year, 0, 1);
      const end = new Date(year, 11, 31, 23, 59, 59);

      const events = cal.getEvents(start, end);

      holidays = events.map(e => ({
        date: Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        title: e.getTitle()
      }));
    }

  } catch (err) {

    // Holiday calendar unreachable (e.g. no network/authorization yet) —
    // fail quietly, calendar still works for events, just without holidays.
    holidays = [];
  }

  cache.put(cacheKey, JSON.stringify(holidays), 21600); // 6 hours

  return holidays;
}

/* =========================
   EMAIL REMINDERS (optional, 2 days before)
   Any logged-in user can subscribe; a daily trigger then emails every
   subscriber about events landing exactly 2 days out.
========================= */

function getReminderSubscribers() {

  const raw = PropertiesService.getScriptProperties().getProperty(REMINDER_SUBS_KEY);

  return raw ? JSON.parse(raw) : [];
}

function isSubscribedToReminders(email) {

  return getReminderSubscribers().indexOf(email) !== -1;
}

function subscribeToReminders(email) {

  if (!email) return { success: false };

  const subs = getReminderSubscribers();

  if (subs.indexOf(email) === -1) {

    subs.push(email);

    PropertiesService.getScriptProperties().setProperty(
      REMINDER_SUBS_KEY,
      JSON.stringify(subs)
    );
  }

  ensureReminderTrigger();

  return { success: true, subscribed: true };
}

function unsubscribeFromReminders(email) {

  const subs = getReminderSubscribers().filter(e => e !== email);

  PropertiesService.getScriptProperties().setProperty(
    REMINDER_SUBS_KEY,
    JSON.stringify(subs)
  );

  return { success: true, subscribed: false };
}

// Make sure the daily reminder trigger exists (created once, on first subscribe).
function ensureReminderTrigger() {

  const triggers = ScriptApp.getProjectTriggers();

  const exists = triggers.some(t => t.getHandlerFunction() === 'sendEventReminders');

  if (!exists) {

    ScriptApp.newTrigger('sendEventReminders')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
  }
}

// Runs daily via the trigger above. Emails every subscriber about
// Scheduled events that fall exactly 2 days from today.
function sendEventReminders() {

  const subs = getReminderSubscribers();

  if (subs.length === 0) return;

  const target = new Date();
  target.setDate(target.getDate() + 2);

  const targetISO = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const events = getEvents().filter(
    e => e.status === 'Scheduled' && e.eventDateISO === targetISO
  );

  if (events.length === 0) return;

  const lines = events.map(e => {

    let line = `• ${e.eventName}${e.company ? ' — ' + e.company : ''}`;

    if (e.type === 'Campus Drive') {
      line += ` (Campus Drive${e.branches ? ': ' + e.branches : ''})`;
    }

    return line;
  });

  const body =
    `Reminder — the following event(s) are coming up on ${targetISO} (2 days from today):\n\n` +
    lines.join('\n') +
    `\n\nOpen the CRM calendar for full details.`;

  subs.forEach(email => {

    try {
      MailApp.sendEmail(email, 'Upcoming Event Reminder — Gyan Ganga Placement Cell', body);
    } catch (err) {
      // Invalid/unreachable address — skip it, don't block the rest.
    }
  });
}