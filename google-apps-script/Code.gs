/**
 * Google Apps Script ללוח השמחות הכיתתי
 *
 * הגדרה:
 * 1. צרי גיליון Google Sheets עם שני טאבים: Events ו-Messages
 * 2. Events – עמודות: id | title | date | time
 * 3. Messages – עמודות: id | author | text | timestamp
 * 4. Extensions → Apps Script → הדביקי את הקוד → Deploy → Web app
 * 5. "Execute as": Me | "Who has access": Anyone
 * 6. העתיקי את כתובת ה-Web App ל-index.html (GOOGLE_SCRIPT_URL)
 */

const EVENTS_SHEET_NAME = 'Events';
const MESSAGES_SHEET_NAME = 'Messages';

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    const params = e ? e.parameter : {};

    switch (action) {
      case 'list':
        return jsonResponse({ success: true, events: listEvents_() });

      case 'add':
        return jsonResponse(addEvent_(params));

      case 'update':
        return jsonResponse(updateEvent_(params));

      case 'delete':
        return jsonResponse(deleteEvent_(params));

      case 'listMessages':
        return jsonResponse({ success: true, messages: listMessages_() });

      case 'addMessage':
        return jsonResponse(addMessage_(params));

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: String(error) });
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getEventsSheet_() {
  return getOrCreateSheet_(EVENTS_SHEET_NAME, ['id', 'title', 'date', 'time']);
}

function getMessagesSheet_() {
  return getOrCreateSheet_(MESSAGES_SHEET_NAME, ['id', 'author', 'text', 'timestamp']);
}

function rowToEvent_(row) {
  return {
    id: Number(row[0]),
    title: String(row[1] || ''),
    date: String(row[2] || ''),
    time: String(row[3] || '19:00')
  };
}

function rowToMessage_(row) {
  return {
    id: Number(row[0]),
    author: String(row[1] || ''),
    text: String(row[2] || ''),
    timestamp: String(row[3] || '')
  };
}

function listEvents_() {
  const sheet = getEventsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter((row) => row[1] && row[2])
    .map(rowToEvent_);
}

function getNextEventId_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return 1;

  const ids = values.slice(1)
    .map((row) => Number(row[0]))
    .filter((id) => !isNaN(id));

  return ids.length ? Math.max.apply(null, ids) + 1 : 1;
}

function addEvent_(params) {
  const title = String(params.title || '').trim();
  const date = String(params.date || '').trim();
  const time = String(params.time || '19:00').trim();

  if (!title || !date) {
    return { success: false, error: 'Missing title or date' };
  }

  const sheet = getEventsSheet_();
  const id = getNextEventId_(sheet);
  sheet.appendRow([id, title, date, time]);

  return { success: true, id: id };
}

function updateEvent_(params) {
  const id = Number(params.id);
  const title = String(params.title || '').trim();
  const date = String(params.date || '').trim();
  const time = String(params.time || '19:00').trim();

  if (!id || !title || !date) {
    return { success: false, error: 'Missing id, title or date' };
  }

  const sheet = getEventsSheet_();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === id) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[title, date, time]]);
      return { success: true, id: id };
    }
  }

  return { success: false, error: 'Event not found' };
}

function deleteEvent_(params) {
  const id = Number(params.id);
  if (!id) {
    return { success: false, error: 'Missing id' };
  }

  const sheet = getEventsSheet_();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }

  return { success: false, error: 'Event not found' };
}

function listMessages_() {
  const sheet = getMessagesSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter((row) => row[2])
    .map(rowToMessage_)
    .sort(function (a, b) {
      return new Date(a.timestamp) - new Date(b.timestamp);
    });
}

function getNextMessageId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  const lastId = Number(sheet.getRange(lastRow, 1).getValue());
  return isNaN(lastId) ? lastRow : lastId + 1;
}

function addMessage_(params) {
  const author = String(params.author || '').trim().slice(0, 40);
  const text = String(params.text || '').trim().slice(0, 500);

  if (!author || !text) {
    return { success: false, error: 'Missing author or text' };
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const sheet = getMessagesSheet_();
    const id = getNextMessageId_(sheet);
    const timestamp = new Date().toISOString();

    sheet.appendRow([id, author, text, timestamp]);

    return {
      success: true,
      id: id,
      timestamp: timestamp,
      message: { id: id, author: author, text: text, timestamp: timestamp }
    };
  } finally {
    lock.releaseLock();
  }
}
