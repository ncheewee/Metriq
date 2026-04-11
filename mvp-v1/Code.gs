/**
 * MetrIQ MVP v1 — Google Apps Script Backend
 * Sheet: Meters | Readings | Users | Config
 *
 * IMPORTANT: This script must be bound to a Google Sheet.
 * Do NOT create it as a standalone script.
 *
 * Correct setup:
 *   1. Open Google Sheets → Extensions → Apps Script
 *   2. Paste this entire file
 *   3. Run setupSheets() once
 *   4. Deploy → New Deployment → Web App
 *
 * Copy the Web App URL into Admin → Settings → Backend URL
 */

// ─── SHEET NAMES ───────────────────────────────────────
const SH_METERS   = 'Meters';
const SH_READINGS = 'Readings';
const SH_USERS    = 'Users';
const SH_CONFIG   = 'Config';

// ─── SPREADSHEET HELPER ─────────────────────────────────
// Always use this instead of SpreadsheetApp.getActiveSpreadsheet()
// Works whether script is bound to a sheet or runs as a web app.
function getSS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script must be bound to a Google Sheet. Open Sheets → Extensions → Apps Script, not script.google.com directly.');
  return ss;
}

// Simple shared API key — stored in Config sheet, key = 'api_key'
// Admin sets this once; it's embedded in the app's admin settings.

// ─── CORS / ENTRY POINT ────────────────────────────────
function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.postData) throw new Error('No POST data received');
    const body = JSON.parse(e.postData.contents);
    const result = route(body);
    out.setContent(JSON.stringify({ ok: true, data: result }));
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return out;
}

function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  // If a 'payload' param is present, treat it as an API call (CORS-safe workaround)
  if (e && e.parameter && e.parameter.payload) {
    try {
      const body = JSON.parse(e.parameter.payload);
      const result = route(body);
      out.setContent(JSON.stringify({ ok: true, data: result }));
    } catch (err) {
      out.setContent(JSON.stringify({ ok: false, error: err.message }));
    }
    return out;
  }

  // Plain GET = health check
  out.setContent(JSON.stringify({ ok: true, version: 'MetrIQ MVP v1' }));
  return out;
}

// ─── ROUTER ────────────────────────────────────────────
function route(body) {
  const { action, apiKey } = body;

  // Keepalive ping — no API key required (used by UptimeRobot to keep VM warm)
  if (action === 'ping') return { pong: true, ts: new Date().toISOString() };

  // Verify API key for all other actions
  if (!verifyKey(apiKey)) throw new Error('Unauthorized');

  switch (action) {
    // METERS
    case 'getMeters':    return getMeters();
    case 'addMeter':     return addMeter(body);
    case 'updateMeter':  return updateMeter(body);
    case 'deleteMeter':  return deleteMeter(body);

    // READINGS
    case 'getReadings':  return getReadings(body);
    case 'addReading':   return addReading(body);
    case 'deleteReading':return deleteReading(body);
    case 'batchSync':    return batchSync(body);

    // USERS
    case 'getUsers':     return getUsers();
    case 'addUser':      return addUser(body);
    case 'updateUser':   return updateUser(body);
    case 'deleteUser':   return deleteUser(body);
    case 'verifyPin':    return verifyPin(body);

    // CONFIG
    case 'getConfig':    return getConfig();
    case 'setConfig':    return setConfig(body);

    // REPORTS
    case 'getMonthSummary': return getMonthSummary(body);
    case 'saveAsCopy':      return saveAsCopy(body);

    // QUOTA TRACKING
    case 'incrementEmailQuota': return incrementEmailQuota();
    case 'resetEmailQuota':     return resetEmailQuota();

    default: throw new Error('Unknown action: ' + action);
  }
}

// ─── KEY VERIFICATION ──────────────────────────────────
function verifyKey(key) {
  const ss = getSS();
  const sheet = ss.getSheetByName(SH_CONFIG);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === 'api_key') return row[1] === key;
  }
  return false;
}

// ─── HELPERS ───────────────────────────────────────────
function getSheet(name) {
  const ss = getSS();
  const s = ss.getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name + '. Run setupSheets() first.');
  return s;
}

function sheetToObjects(sheet) {
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(sheet, headers, obj) {
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.appendRow(row);
}

function findRowIndex(sheet, colIndex, value) {
  // Returns 1-based row index (including header), or -1
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

function nowISO() {
  return new Date().toISOString();
}

// ─── METERS ────────────────────────────────────────────
const METER_HEADERS = ['meter_id','label','location','floor','active','added_date','added_by'];

function getMeters() {
  const sheet = getSheet(SH_METERS);
  const rows = sheetToObjects(sheet);
  return rows.filter(r => r.active !== 'N');
}

function addMeter(body) {
  const { label, location, floor, addedBy } = body;
  if (!label) throw new Error('label is required');
  const sheet = getSheet(SH_METERS);
  const meter = {
    meter_id: 'M_' + Date.now(),
    label: label,
    location: location || '',
    floor: floor || '',
    active: 'Y',
    added_date: nowISO(),
    added_by: addedBy || ''
  };
  appendRow(sheet, METER_HEADERS, meter);
  return meter;
}

function updateMeter(body) {
  const { meter_id, label, location, floor } = body;
  const sheet = getSheet(SH_METERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex(sheet, headers.indexOf('meter_id'), meter_id);
  if (rowIdx < 0) throw new Error('Meter not found');
  const updates = { label, location, floor };
  Object.entries(updates).forEach(([k, v]) => {
    if (v !== undefined) {
      const col = headers.indexOf(k) + 1;
      if (col > 0) sheet.getRange(rowIdx, col).setValue(v);
    }
  });
  return { ok: true };
}

function deleteMeter(body) {
  const { meter_id } = body;
  const sheet = getSheet(SH_METERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex(sheet, headers.indexOf('meter_id'), meter_id);
  if (rowIdx < 0) throw new Error('Meter not found');
  const col = headers.indexOf('active') + 1;
  sheet.getRange(rowIdx, col).setValue('N'); // soft delete
  return { ok: true };
}

// ─── READINGS ──────────────────────────────────────────
const READING_HEADERS = [
  'reading_id','meter_id','month_year','reading_value','reading_date',
  'reading_time','inspector_id','inspector_name','photo_url','ai_confidence',
  'ai_provider','notes','synced_at'
];

function getReadings(body) {
  const { month_year, meter_id } = body;
  const sheet = getSheet(SH_READINGS);
  let rows = sheetToObjects(sheet);
  if (month_year) rows = rows.filter(r => r.month_year === month_year);
  if (meter_id) rows = rows.filter(r => r.meter_id === meter_id);
  return rows;
}

function addReading(body) {
  const { meter_id, month_year, reading_value, reading_date, reading_time,
          inspector_id, inspector_name, photo_b64, ai_confidence, ai_provider, notes } = body;

  if (!meter_id || !month_year || reading_value === undefined)
    throw new Error('meter_id, month_year, reading_value required');

  // Check: only 1 reading per meter per calendar month
  const sheet = getSheet(SH_READINGS);
  const existing = sheetToObjects(sheet).filter(
    r => r.meter_id === meter_id && r.month_year === month_year
  );
  if (existing.length > 0) throw new Error('Reading already exists for this meter in ' + month_year);

  // Upload photo to Drive if provided
  let photo_url = '';
  if (photo_b64) {
    photo_url = uploadPhotoToDrive(meter_id, month_year, photo_b64);
  }

  const reading = {
    reading_id: 'R_' + Date.now(),
    meter_id,
    month_year,
    reading_value: String(reading_value),
    reading_date: reading_date || '',
    reading_time: reading_time || '',
    inspector_id: inspector_id || '',
    inspector_name: inspector_name || '',
    photo_url,
    ai_confidence: ai_confidence || '',
    ai_provider: ai_provider || '',
    notes: notes || '',
    synced_at: nowISO()
  };
  appendRow(sheet, READING_HEADERS, reading);
  return reading;
}

function deleteReading(body) {
  const { reading_id } = body;
  const sheet = getSheet(SH_READINGS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex(sheet, headers.indexOf('reading_id'), reading_id);
  if (rowIdx < 0) throw new Error('Reading not found');
  sheet.deleteRow(rowIdx);
  return { ok: true };
}

// Sync a batch of readings captured offline
function batchSync(body) {
  const { readings } = body; // array of reading objects
  if (!Array.isArray(readings)) throw new Error('readings must be an array');
  const results = [];
  for (const r of readings) {
    try {
      const result = addReading(r);
      results.push({ reading_id: r.reading_id || r.meter_id, ok: true, server_id: result.reading_id });
    } catch (err) {
      results.push({ reading_id: r.reading_id || r.meter_id, ok: false, error: err.message });
    }
  }
  return results;
}

// ─── PHOTO UPLOAD ──────────────────────────────────────
function uploadPhotoToDrive(meter_id, month_year, b64) {
  try {
    const config = getConfigMap();
    const folderId = config['drive_folder_id'];

    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      // Create folder on first use, save its ID
      folder = DriveApp.createFolder('MetrIQ Photos');
      setConfigValue('drive_folder_id', folder.getId());
    }

    // Sub-folder per month
    const monthFolderName = 'MetrIQ_' + month_year;
    let monthFolder;
    const iter = folder.getFoldersByName(monthFolderName);
    if (iter.hasNext()) {
      monthFolder = iter.next();
    } else {
      monthFolder = folder.createFolder(monthFolderName);
    }

    // Decode base64
    const parts = b64.split(',');
    const mimeType = parts[0].match(/:(.*?);/)[1];
    const data = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(data, mimeType, meter_id + '_' + month_year + '.jpg');

    const file = monthFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/file/d/' + file.getId() + '/view';
  } catch (e) {
    console.error('Drive upload failed: ' + e.message);
    return ''; // non-fatal
  }
}

// ─── USERS ─────────────────────────────────────────────
const USER_HEADERS = ['user_id','name','role','pin_hash','telegram_chat_id','email','active','created_at'];

function getUsers() {
  const sheet = getSheet(SH_USERS);
  const rows = sheetToObjects(sheet);
  // Never expose pin_hash to client — strip it
  return rows
    .filter(r => r.active !== 'N')
    .map(r => ({ ...r, pin_hash: undefined }));
}

function addUser(body) {
  const { name, role, pin, telegram_chat_id, email, createdBy } = body;
  if (!name || !role || !pin) throw new Error('name, role, pin required');
  if (!['Inspector','Manager','Admin'].includes(role)) throw new Error('Invalid role');
  if (!/^\d{6}$/.test(pin)) throw new Error('PIN must be 6 digits');

  const sheet = getSheet(SH_USERS);
  const user = {
    user_id: 'U_' + Date.now(),
    name,
    role,
    pin_hash: hashPin(pin),
    telegram_chat_id: telegram_chat_id || '',
    email: email || '',
    active: 'Y',
    created_at: nowISO()
  };
  appendRow(sheet, USER_HEADERS, user);
  const { pin_hash, ...safe } = user;
  return safe;
}

function updateUser(body) {
  const { user_id, name, telegram_chat_id, email, pin } = body;
  const sheet = getSheet(SH_USERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex(sheet, headers.indexOf('user_id'), user_id);
  if (rowIdx < 0) throw new Error('User not found');

  const updates = { name, telegram_chat_id, email };
  if (pin) {
    if (!/^\d{6}$/.test(pin)) throw new Error('PIN must be 6 digits');
    updates.pin_hash = hashPin(pin);
  }
  Object.entries(updates).forEach(([k, v]) => {
    if (v !== undefined) {
      const col = headers.indexOf(k) + 1;
      if (col > 0) sheet.getRange(rowIdx, col).setValue(v);
    }
  });
  return { ok: true };
}

function deleteUser(body) {
  const { user_id } = body;
  const sheet = getSheet(SH_USERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex(sheet, headers.indexOf('user_id'), user_id);
  if (rowIdx < 0) throw new Error('User not found');
  const col = headers.indexOf('active') + 1;
  sheet.getRange(rowIdx, col).setValue('N');
  return { ok: true };
}

function verifyPin(body) {
  const { role, pin } = body;
  if (!role || !pin) throw new Error('role and pin required');
  if (!/^\d{6}$/.test(pin)) return { valid: false };

  const sheet = getSheet(SH_USERS);
  const rows = sheetToObjects(sheet);
  const hash = hashPin(pin);

  // Find any active user matching role + pin
  const user = rows.find(r =>
    r.active !== 'N' &&
    r.role === role &&
    r.pin_hash === hash
  );

  if (user) {
    const { pin_hash, ...safe } = user;
    return { valid: true, user: safe };
  }
  return { valid: false };
}

// Simple SHA-256 equivalent using Apps Script Utilities
function hashPin(pin) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pin + '_metriq_salt'
  );
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// ─── CONFIG ────────────────────────────────────────────
function getConfigMap() {
  const sheet = getSheet(SH_CONFIG);
  const data = sheet.getDataRange().getValues();
  const map = {};
  data.forEach(row => { if (row[0]) map[String(row[0])] = row[1]; });
  return map;
}

function getConfig() {
  const map = getConfigMap();
  // Don't expose api_key to client
  const { api_key, ...safe } = map;
  return safe;
}

function setConfig(body) {
  const { updates } = body; // { key: value, ... }
  if (!updates) throw new Error('updates required');
  for (const [k, v] of Object.entries(updates)) {
    if (k === 'api_key') continue; // never allow overwrite via client
    setConfigValue(k, v);
  }
  return { ok: true };
}

function setConfigValue(key, value) {
  const sheet = getSheet(SH_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ─── MONTHLY SUMMARY ───────────────────────────────────
function getMonthSummary(body) {
  const { month_year } = body;
  if (!month_year) throw new Error('month_year required (e.g. 2026-04)');

  const meters = getMeters();
  const readings = getReadings({ month_year });

  const readingMap = {};
  readings.forEach(r => { readingMap[r.meter_id] = r; });

  const summary = meters.map(m => ({
    meter_id: m.meter_id,
    label: m.label,
    location: m.location,
    floor: m.floor,
    reading: readingMap[m.meter_id] ? readingMap[m.meter_id].reading_value : null,
    reading_date: readingMap[m.meter_id] ? readingMap[m.meter_id].reading_date : null,
    inspector: readingMap[m.meter_id] ? readingMap[m.meter_id].inspector_name : null,
    photo_url: readingMap[m.meter_id] ? readingMap[m.meter_id].photo_url : null,
    done: !!readingMap[m.meter_id]
  }));

  const done = summary.filter(s => s.done).length;
  const pending = summary.filter(s => !s.done).map(s => s.label);

  return {
    month_year,
    total_meters: meters.length,
    done,
    pending_count: meters.length - done,
    completion_pct: meters.length > 0 ? Math.round((done / meters.length) * 100) : 0,
    pending_meters: pending,
    meters: summary
  };
}

// ─── SAVE AS COPY ──────────────────────────────────────
function saveAsCopy(body) {
  const { label } = body;
  const ss = getSS();
  const copyName = (label || 'MetrIQ Data') + ' — Copy ' + Utilities.formatDate(new Date(), 'Asia/Singapore', 'yyyy-MM-dd HH:mm');
  const copy = ss.copy(copyName);
  // Make it accessible to anyone with link
  DriveApp.getFileById(copy.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/edit', name: copyName };
}

// ─── EMAIL QUOTA ───────────────────────────────────────
function incrementEmailQuota() {
  const config = getConfigMap();
  const current = parseInt(config['email_quota_used'] || '0', 10);
  const month = Utilities.formatDate(new Date(), 'Asia/Singapore', 'yyyy-MM');
  const lastReset = config['email_quota_month'] || '';

  // Auto-reset on new month
  if (lastReset !== month) {
    setConfigValue('email_quota_used', 1);
    setConfigValue('email_quota_month', month);
    return { used: 1, limit: 200, month };
  }

  const newCount = current + 1;
  setConfigValue('email_quota_used', newCount);
  return { used: newCount, limit: 200, month };
}

function resetEmailQuota() {
  const month = Utilities.formatDate(new Date(), 'Asia/Singapore', 'yyyy-MM');
  setConfigValue('email_quota_used', 0);
  setConfigValue('email_quota_month', month);
  return { ok: true };
}

// ─── SETUP HELPER ──────────────────────────────────────
// Run this ONCE manually to set up sheet headers and default config.
//
// HOW TO RUN:
//   In Apps Script editor → select 'setupSheets' from the function dropdown → click Run
//
function setupSheets() {
  const ss = getSS();

  function ensureSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    // Only write headers if sheet is empty
    if (s.getLastRow() === 0) {
      s.appendRow(headers);
      s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
      s.setFrozenRows(1);
    }
    return s;
  }

  ensureSheet(SH_METERS,   METER_HEADERS);
  ensureSheet(SH_READINGS, READING_HEADERS);
  ensureSheet(SH_USERS,    USER_HEADERS);
  ensureSheet(SH_CONFIG,   ['key', 'value']);

  // Remove the default blank 'Sheet1' if it exists
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Default config values
  const configSheet = ss.getSheetByName(SH_CONFIG);
  const existingData = configSheet.getDataRange().getValues();
  const existingKeys = existingData.map(r => r[0]);

  const defaults = [
    ['api_key',            'METRIQ_' + Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase()],
    ['site_name',          'My Site'],
    ['email_quota_used',   '0'],
    ['email_quota_month',  ''],
    ['drive_folder_id',    ''],
    ['telegram_token',     ''],
    ['telegram_chat_id',   ''],
    ['emailjs_service_id', ''],
    ['emailjs_template_id',''],
    ['emailjs_public_key', ''],
  ];

  defaults.forEach(([k, v]) => {
    if (!existingKeys.includes(k)) configSheet.appendRow([k, v]);
  });

  // Protect the Config sheet (warning only — prevents accidental edits)
  const protection = configSheet.protect().setDescription('Config — Admin only');
  protection.setWarningOnly(true);

  // Show the api_key in a dialog so you can copy it immediately
  const apiKey = configSheet.getDataRange().getValues().find(r => r[0] === 'api_key')?.[1] || '(see Config sheet)';

  SpreadsheetApp.getUi().alert(
    '✅ MetrIQ Setup Complete!\n\n' +
    'API Key (copy this now):\n' + apiKey + '\n\n' +
    'Steps:\n' +
    '1. Copy the API key above\n' +
    '2. Deploy → New Deployment → Web App\n' +
    '   Execute as: Me | Access: Anyone\n' +
    '3. Copy the Web App URL\n' +
    '4. In the app: Admin → Settings → paste URL + API key\n\n' +
    'The API key is also in the Config sheet (api_key row).'
  );
}

// ─── ADD FIRST ADMIN ──────────────────────────────────
// Run this ONCE after setupSheets() to create your first Admin user.
// Change the values below before running.
function addFirstAdmin() {
  const NAME  = 'Admin';        // ← change to your name
  const PIN   = '123456';       // ← change to your 6-digit PIN
  const EMAIL = 'you@email.com'; // ← change to your email

  const sheet = getSheet(SH_USERS);
  const existing = sheetToObjects(sheet);
  if (existing.some(u => u.role === 'Admin' && u.active !== 'N')) {
    SpreadsheetApp.getUi().alert('An active Admin user already exists. Use the app to add more users.');
    return;
  }
  if (!/^\d{6}$/.test(PIN)) {
    SpreadsheetApp.getUi().alert('PIN must be exactly 6 digits.');
    return;
  }
  const user = {
    user_id: 'U_' + Date.now(),
    name: NAME,
    role: 'Admin',
    pin_hash: hashPin(PIN),
    telegram_chat_id: '',
    email: EMAIL,
    active: 'Y',
    created_at: new Date().toISOString()
  };
  appendRow(sheet, USER_HEADERS, user);
  SpreadsheetApp.getUi().alert(
    '✅ Admin user created!\n\n' +
    'Name: ' + NAME + '\n' +
    'PIN: ' + PIN + '\n\n' +
    'You can now log into the app as Admin.'
  );
}
