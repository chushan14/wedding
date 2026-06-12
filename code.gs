const SPREADSHEET_ID = '1V_mSJSzTk6V9dSi_7r1CXOuXCROANhxMv5z7ZZtnobw';

const SHEETS = {
  guests: 'Guests',
  gifts: 'Gifts',
  settings: 'Settings'
};

const HEADERS = {
  Guests: [
    'id', 'name', 'phone', 'relation', 'invitation', 'invitationEmail',
    'invitationAddress', 'attendance', 'count', 'companionNames',
    'mealsJson', 'babyChairs', 'selfDrive', 'carPlate', 'wish', 'tableId',
    'checkedIn', 'deletedAt', 'createdAt', 'updatedAt'
  ],
  Gifts: ['id', 'giver', 'amount', 'note', 'deletedAt', 'createdAt', 'updatedAt'],
  Settings: ['key', 'value', 'updatedAt']
};

function doGet() {
  return json({
    ok: true,
    app: '12 hours closer RSVP API',
    message: 'API is running. Deploy GitHub Pages for the frontend.'
  });
}

function doPost(e) {
  try {
    ensureSheets();
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = body.action;

    if (action === 'createGuest') return json(createGuest(body.guest));
    if (action === 'listGuests') return json(listGuests());
    if (action === 'deleteGuest') return json(deleteGuest(body.id));
    if (action === 'updateGuest') return json(updateGuest(body));

    if (action === 'createGift') return json(createGift(body.gift));
    if (action === 'listGifts') return json(listGifts());
    if (action === 'deleteGift') return json(deleteGift(body.id));

    if (action === 'getSettings') return json(getSettings());
    if (action === 'saveSetting') return json(saveSetting(body));

    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (error) {
    return json({ ok: false, error: error.message || String(error) });
  }
}

function createGuest(guest) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet(SHEETS.guests);
    const now = new Date().toISOString();
    const id = guest.id || String(Date.now());

    appendObject(sheet, {
      id,
      name: guest.name || '',
      phone: guest.phone || '',
      relation: guest.relation || '',
      invitation: guest.invitation || '',
      invitationEmail: guest.invitationEmail || '',
      invitationAddress: guest.invitationAddress || '',
      attendance: guest.attendance || '',
      count: Number(guest.count || 0),
      companionNames: guest.companionNames || '',
      mealsJson: JSON.stringify(guest.meals || []),
      babyChairs: Number(guest.babyChairs || 0),
      selfDrive: guest.selfDrive || 'no',
      carPlate: guest.carPlate || '',
      wish: guest.wish || '',
      tableId: guest.tableId || '',
      checkedIn: guest.checkedIn === true,
      deletedAt: '',
      createdAt: now,
      updatedAt: now
    });

    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function listGuests() {
  const rows = readObjects(SHEETS.guests);
  const guests = rows
    .filter(row => !row.deletedAt)
    .map(row => {
      try {
        row.meals = JSON.parse(row.mealsJson || '[]');
      } catch (e) {
        row.meals = [];
      }
      return row;
    });

  return { ok: true, guests };
}

function deleteGuest(id) {
  return updateRowById(SHEETS.guests, id, {
    deletedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });
}

function updateGuest(payload) {
  const allowed = [
    'name', 'phone', 'relation', 'invitation', 'invitationEmail',
    'invitationAddress', 'attendance', 'count', 'companionNames', 'mealsJson',
    'babyChairs', 'selfDrive', 'carPlate', 'wish', 'tableId', 'checkedIn'
  ];
  const fields = {};
  Object.keys(payload.fields || {}).forEach(key => {
    if (allowed.includes(key)) fields[key] = payload.fields[key];
  });
  fields.updatedAt = new Date().toISOString();
  return updateRowById(SHEETS.guests, payload.id, fields);
}

function createGift(gift) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet(SHEETS.gifts);
    const now = new Date().toISOString();
    const id = gift.id || 'gift-' + Date.now();

    appendObject(sheet, {
      id,
      giver: gift.giver || '',
      amount: Number(gift.amount || 0),
      note: gift.note || '',
      deletedAt: '',
      createdAt: now,
      updatedAt: now
    });

    markGuestCheckedInByName(gift.giver);
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function listGifts() {
  const gifts = readObjects(SHEETS.gifts)
    .filter(row => !row.deletedAt)
    .map(row => ({
      id: row.id,
      giver: row.giver || '',
      amount: Number(row.amount || 0),
      note: row.note || ''
    }));

  return { ok: true, gifts };
}

function deleteGift(id) {
  return updateRowById(SHEETS.gifts, id, {
    deletedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });
}

function getSettings() {
  const settings = {};
  readObjects(SHEETS.settings).forEach(row => {
    if (row.key) settings[row.key] = row.value || '';
  });
  return { ok: true, settings };
}

function saveSetting(payload) {
  const key = payload.key;
  const value = payload.value || '';
  if (!key) return { ok: false, error: 'Missing setting key' };

  const sheet = getSheet(SHEETS.settings);
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();

  for (let r = 2; r <= data.length; r++) {
    if (String(sheet.getRange(r, 1).getValue()) === String(key)) {
      sheet.getRange(r, 2).setValue(value);
      sheet.getRange(r, 3).setValue(now);
      return { ok: true };
    }
  }

  sheet.appendRow([key, value, now]);
  return { ok: true };
}

function markGuestCheckedInByName(name) {
  if (!name) return;
  const sheet = getSheet(SHEETS.guests);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('name') + 1;
  const attendanceCol = headers.indexOf('attendance') + 1;
  const checkedInCol = headers.indexOf('checkedIn') + 1;
  const deletedAtCol = headers.indexOf('deletedAt') + 1;
  const updatedAtCol = headers.indexOf('updatedAt') + 1;

  for (let r = 2; r <= data.length; r++) {
    const rowName = String(sheet.getRange(r, nameCol).getValue()).trim();
    const attendance = String(sheet.getRange(r, attendanceCol).getValue()).trim();
    const deletedAt = String(sheet.getRange(r, deletedAtCol).getValue()).trim();
    if (rowName === String(name).trim() && attendance === 'yes' && !deletedAt) {
      sheet.getRange(r, checkedInCol).setValue(true);
      sheet.getRange(r, updatedAtCol).setValue(new Date().toISOString());
      return;
    }
  }
}

function updateRowById(sheetName, id, fields) {
  if (!id) return { ok: false, error: 'Missing id' };

  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id') + 1;

  for (let r = 2; r <= data.length; r++) {
    if (String(sheet.getRange(r, idCol).getValue()) === String(id)) {
      Object.keys(fields).forEach(key => {
        const col = headers.indexOf(key) + 1;
        if (col > 0) sheet.getRange(r, col).setValue(fields[key]);
      });
      return { ok: true };
    }
  }

  return { ok: false, error: 'Row not found: ' + id };
}

function readObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  return values.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

function appendObject(sheet, object) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '');
  sheet.appendRow(row);
}

function getSheet(sheetName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
}

function ensureSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(HEADERS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    const expected = HEADERS[sheetName];
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
      return;
    }

    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(Boolean);
    if (current.length === 0) {
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
      return;
    }

    expected.forEach(header => {
      if (!current.includes(header)) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      }
    });
  });
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
