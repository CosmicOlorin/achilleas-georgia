const SHEET_ID = '1Rba07Brfx8DG550CWlS0hwg1kLTSzwClKt237jD1Tds';
const RESPONSE_HEADERS = ['Ημερομηνία', 'Visitor ID', 'Ονοματεπώνυμο', 'Τηλέφωνο', 'RSVP', 'Άτομα', 'Μήνυμα', 'Επιλογή', 'Ποσό', 'Τραπέζι', 'Τελευταία ενημέρωση'];
const WISH_HEADERS = ['Ημερομηνία', 'Visitor ID', 'Ονοματεπώνυμο', 'Τηλέφωνο RSVP', 'Ευχή'];

function doGet() {
  return ContentService.createTextOutput('Wedding endpoint is active');
}

function ensureSheet_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function migrateResponses_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Απαντήσεις') || spreadsheet.insertSheet('Απαντήσεις');
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const oldHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  if (oldHeaders.join('|') === RESPONSE_HEADERS.join('|')) return sheet;

  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
  const column = name => oldHeaders.indexOf(name);
  const value = (row, name) => column(name) >= 0 ? row[column(name)] : '';
  const migrated = rows.map(row => {
    let choice = value(row, 'Επιλογή');
    if (!choice) {
      if (value(row, 'Τελευταίο κλικ') === 'IBAN' || value(row, 'IBAN') === 'Ναι') choice = 'IBAN';
      else if (value(row, 'Τελευταίο κλικ') === 'IRIS' || value(row, 'IRIS') === 'Ναι') choice = 'IRIS';
      else if (value(row, 'Τελευταίο κλικ') === 'ΟΧΙ' || value(row, 'Δώρο') === 'Όχι') choice = 'ΚΑΝΕΝΑ ΔΩΡΟ';
      else choice = 'ΚΑΜΙΑ ΕΠΙΛΟΓΗ';
    }
    return [
      value(row, 'Ημερομηνία'), value(row, 'Visitor ID'), value(row, 'Ονοματεπώνυμο'),
      value(row, 'Τηλέφωνο'), value(row, 'RSVP'), value(row, 'Άτομα'), value(row, 'Μήνυμα'),
      choice, value(row, 'Ποσό'), value(row, 'Τραπέζι'), value(row, 'Τελευταία ενημέρωση')
    ];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
  if (migrated.length) sheet.getRange(2, 1, migrated.length, RESPONSE_HEADERS.length).setValues(migrated);
  sheet.setFrozenRows(1);
  return sheet;
}

function responseRow_(sheet, visitorId) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 2, last - 1, 1).getDisplayValues().flat();
  const index = ids.lastIndexOf(visitorId);
  return index < 0 ? 0 : index + 2;
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function migrateGiftColumns() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  migrateResponses_(spreadsheet);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const responses = migrateResponses_(spreadsheet);
    const visitorId = p.visitorId || Utilities.getUuid();
    let row = responseRow_(responses, visitorId);

    if (p.action === 'wish') {
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

    if (!row) {
      responses.appendRow([new Date(), visitorId, '', '', '', '', '', 'ΚΑΜΙΑ ΕΠΙΛΟΓΗ', '', '', new Date()]);
      row = responses.getLastRow();
    }

    if (p.action === 'rsvp') {
      responses.getRange(row, 3, 1, 5).setValues([[p.name || '', p.phone || '', p.attendance || '', p.guests || '', p.message || '']]);
    }

    if (p.action === 'gift_click') {
      const choices = { IBAN: 'IBAN', IRIS: 'IRIS', 'ΟΧΙ': 'ΚΑΝΕΝΑ ΔΩΡΟ' };
      responses.getRange(row, 8).setValue(choices[p.gift] || 'ΚΑΜΙΑ ΕΠΙΛΟΓΗ');
    }

    // Columns 9 (Ποσό) and 10 (Τραπέζι) are intentionally never changed by the site.
    responses.getRange(row, 11).setValue(new Date());
    return json_({ ok: true, sheet: 'Απαντήσεις' });
  } finally {
    lock.releaseLock();
  }
}
