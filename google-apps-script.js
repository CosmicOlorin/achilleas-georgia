const SHEET_ID = '1Rba07Brfx8DG550CWlS0hwg1kLTSzwClKt237jD1Tds';
const RESPONSE_HEADERS = ['Ημερομηνία', 'Visitor ID', 'Ονοματεπώνυμο', 'Τηλέφωνο', 'RSVP', 'Άτομα', 'Μήνυμα', 'Επιλογή', 'Ποσό', 'Τραπέζι', 'Τελευταία ενημέρωση'];
const WISH_HEADERS = ['Ημερομηνία', 'Visitor ID', 'Ονοματεπώνυμο', 'Τηλέφωνο RSVP', 'Ευχή'];

function doGet() {
  return ContentService.createTextOutput('Wedding endpoint is active');
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheet_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function responsesSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Απαντήσεις');
  if (!sheet) throw new Error('Δεν βρέθηκε η καρτέλα Απαντήσεις.');
  const headers = sheet.getRange(1, 1, 1, 11).getDisplayValues()[0];
  if (headers.join('|') !== RESPONSE_HEADERS.join('|')) {
    throw new Error('Η δομή της καρτέλας Απαντήσεις έχει αλλάξει. Δεν έγινε καμία εγγραφή.');
  }
  return sheet;
}

function responseRow_(sheet, visitorId) {
  const ids = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).getDisplayValues().flat();
  const index = ids.lastIndexOf(visitorId);
  return index < 0 ? 0 : index + 2;
}

function firstEmptyResponseRow_(sheet) {
  const ids = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).getDisplayValues().flat();
  const index = ids.findIndex(value => !value);
  if (index >= 0) return index + 2;
  sheet.insertRowAfter(sheet.getMaxRows());
  return sheet.getMaxRows();
}

function sameValues_(current, next) {
  return current.length === next.length && current.every((value, index) => String(value) === String(next[index]));
}

function markChanged_(sheet, row) {
  sheet.getRange(row, 11).setValue(new Date());
  sheet.getRange(row, 12).setValue(false);
}

function createResponse_(sheet, visitorId, values) {
  const row = firstEmptyResponseRow_(sheet);
  const now = new Date();
  sheet.getRange(row, 1, 1, 12).setValues([[
    now, visitorId, values.name || '', values.phone || '', values.attendance || '',
    values.guests || '', values.message || '', values.choice || 'ΚΑΜΙΑ ΕΠΙΛΟΓΗ',
    '', '', now, false
  ]]);
  return row;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || '';
    if (!['rsvp', 'gift_click', 'wish'].includes(action)) return json_({ ok: true, ignored: true });

    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const responses = responsesSheet_(spreadsheet);
    const visitorId = p.visitorId || Utilities.getUuid();
    let row = responseRow_(responses, visitorId);

    if (action === 'wish') {
      let rsvpName = '';
      let rsvpPhone = '';
      if (row) {
        rsvpName = responses.getRange(row, 3).getDisplayValue();
        rsvpPhone = responses.getRange(row, 4).getDisplayValue();
      }
      const wishes = ensureSheet_(spreadsheet, 'Ευχολόγιο', WISH_HEADERS);
      wishes.appendRow([new Date(), visitorId, p.wishName || rsvpName, rsvpPhone, p.wish || '']);
      return json_({ ok: true, sheet: 'Ευχολόγιο' });
    }

    if (action === 'rsvp') {
      const next = [p.name || '', p.phone || '', p.attendance || '', p.guests || '', p.message || ''];
      if (!row) {
        createResponse_(responses, visitorId, { name: next[0], phone: next[1], attendance: next[2], guests: next[3], message: next[4] });
        return json_({ ok: true, changed: true, sheet: 'Απαντήσεις' });
      }
      const current = responses.getRange(row, 3, 1, 5).getDisplayValues()[0];
      if (!sameValues_(current, next)) {
        responses.getRange(row, 3, 1, 5).setValues([next]);
        markChanged_(responses, row);
        return json_({ ok: true, changed: true, sheet: 'Απαντήσεις' });
      }
      return json_({ ok: true, changed: false, sheet: 'Απαντήσεις' });
    }

    const choices = { IBAN: 'IBAN', IRIS: 'IRIS', 'ΟΧΙ': 'ΚΑΝΕΝΑ ΔΩΡΟ' };
    const choice = choices[p.gift];
    if (!choice) return json_({ ok: true, ignored: true });
    if (!row) {
      createResponse_(responses, visitorId, { choice: choice });
      return json_({ ok: true, changed: true, sheet: 'Απαντήσεις' });
    }
    const currentChoice = responses.getRange(row, 8).getDisplayValue();
    if (currentChoice !== choice) {
      responses.getRange(row, 8).setValue(choice);
      markChanged_(responses, row);
      return json_({ ok: true, changed: true, sheet: 'Απαντήσεις' });
    }
    return json_({ ok: true, changed: false, sheet: 'Απαντήσεις' });
  } finally {
    lock.releaseLock();
  }
}
