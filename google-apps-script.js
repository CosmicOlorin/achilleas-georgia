const SHEET_ID = '1Rba07Brfx8DG550CWlS0hwg1kLTSzwClKt237jD1Tds';
const RESPONSE_HEADERS = ['Ημερομηνία', 'Visitor ID', 'Ονοματεπώνυμο', 'Τηλέφωνο', 'RSVP', 'Άτομα', 'Μήνυμα', 'Δώρο', 'IBAN', 'IRIS', 'Τελευταίο κλικ', 'Τελευταία ενημέρωση'];
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

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const responses = ensureSheet_(spreadsheet, 'Απαντήσεις', RESPONSE_HEADERS);
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
      responses.appendRow([new Date(), visitorId, '', '', '', '', '', 'Όχι', 'Όχι', 'Όχι', '', new Date()]);
      row = responses.getLastRow();
    }

    if (p.action === 'rsvp') {
      responses.getRange(row, 3, 1, 5).setValues([[p.name || '', p.phone || '', p.attendance || '', p.guests || '', p.message || '']]);
    }

    if (p.action === 'gift_click') {
      const gift = p.gift || '';
      responses.getRange(row, 8).setValue(gift === 'ΟΧΙ' ? 'Όχι' : 'Ναι');
      if (gift === 'IBAN') responses.getRange(row, 9).setValue('Ναι');
      if (gift === 'IRIS') responses.getRange(row, 10).setValue('Ναι');
      responses.getRange(row, 11).setValue(gift);
    }

    responses.getRange(row, 12).setValue(new Date());
    return json_({ ok: true, sheet: 'Απαντήσεις' });
  } finally {
    lock.releaseLock();
  }
}
