const SHARED_SECRET = "sih2026-change-this-secret";

const COOLDOWN_SECONDS = 10 * 60;

const COMPLAINTS_SHEET = 'Complaints';
const ADMINS_SHEET = 'Admins';

const COL = { // Complaints sheet
  TIMESTAMP: 1, TICKET: 2, TEAM_NAME: 3, TEAM_NO: 4, CATEGORY: 5,
  SUBJECT: 6, ISSUE: 7, VENUE: 8, STATUS: 9, REMARKS: 10, UPDATED: 11, HANDLED_BY: 12
};

const ACOL = { // Admins sheet
  NAME: 1, KEY: 2, ROLE: 3, ACTIVE: 4, ADDED_BY: 5, ADDED_ON: 6
};

function complaintsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPLAINTS_SHEET);
}
function adminsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMINS_SHEET);
}

// ---------- Complaint submission (called from index.html) ----------

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.secret !== SHARED_SECRET) return jsonResponse({ result: 'unauthorized' });
  if (data.website) return jsonResponse({ result: 'rejected' }); // honeypot

  const cache = CacheService.getScriptCache();
  const deviceId = data.deviceId || 'unknown';
  if (cache.get(deviceId)) return jsonResponse({ result: 'cooldown' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = complaintsSheet();
    const ticketNum = sheet.getLastRow();
    const ticketId = 'SIH-' + String(ticketNum).padStart(4, '0');

    sheet.appendRow([
      new Date(), ticketId, data.teamName || '', data.teamNo || '', data.category || '',
      data.subject || '', data.issue || '', data.venue || '', 'Pending', '', new Date(), ''
    ]);

    cache.put(deviceId, '1', COOLDOWN_SECONDS);
    return jsonResponse({ result: 'success', ticketId: ticketId });
  } finally {
    lock.releaseLock();
  }
}

// ---------- GET router ----------

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'status') return getStatus(e.parameter.ticketId, e.parameter.teamNo);

  if (action === 'adminLogin') return adminLogin(e.parameter.key);

  if (action === 'list') {
    const admin = requireAdmin(e.parameter.key);
    if (!admin) return jsonResponse({ error: 'unauthorized' });
    return listComplaints();
  }

  if (action === 'update') {
    const admin = requireAdmin(e.parameter.key);
    if (!admin) return jsonResponse({ error: 'unauthorized' });
    return updateComplaint(e.parameter, admin.name);
  }

  if (action === 'listAdmins') {
    const admin = requireOwner(e.parameter.key);
    if (!admin) return jsonResponse({ error: 'unauthorized' });
    return listAdmins();
  }

  if (action === 'addAdmin') {
    const admin = requireOwner(e.parameter.key);
    if (!admin) return jsonResponse({ error: 'unauthorized' });
    return addAdmin(e.parameter.newName, admin.name);
  }

  if (action === 'setAdminActive') {
    const admin = requireOwner(e.parameter.key);
    if (!admin) return jsonResponse({ error: 'unauthorized' });
    return setAdminActive(e.parameter.targetKey, e.parameter.active === 'true');
  }

  return jsonResponse({ error: 'unknown action' });
}

// ---------- Admin identity ----------

function findAdminByKey(key) {
  if (!key) return null;
  const sheet = adminsSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[ACOL.KEY - 1]).trim() === String(key).trim()) {
      return {
        row: i + 1,
        name: row[ACOL.NAME - 1],
        role: row[ACOL.ROLE - 1],
        active: String(row[ACOL.ACTIVE - 1]).toLowerCase() === 'yes'
      };
    }
  }
  return null;
}

function requireAdmin(key) {
  const admin = findAdminByKey(key);
  if (!admin || !admin.active) return null;
  return admin;
}

function requireOwner(key) {
  const admin = requireAdmin(key);
  if (!admin || admin.role !== 'Owner') return null;
  return admin;
}

function adminLogin(key) {
  const admin = requireAdmin(key);
  if (!admin) return jsonResponse({ valid: false });
  return jsonResponse({ valid: true, name: admin.name, role: admin.role });
}

function listAdmins() {
  const sheet = adminsSheet();
  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[ACOL.NAME - 1]) continue;
    const fullKey = String(row[ACOL.KEY - 1]);
    out.push({
      name: row[ACOL.NAME - 1],
      maskedKey: fullKey.length > 4 ? '••••' + fullKey.slice(-4) : '••••',
      key: fullKey, // needed so the owner can revoke by exact key; admin.html never displays this in full for non-owner views
      role: row[ACOL.ROLE - 1],
      active: String(row[ACOL.ACTIVE - 1]).toLowerCase() === 'yes',
      addedBy: row[ACOL.ADDED_BY - 1],
      addedOn: row[ACOL.ADDED_ON - 1]
    });
  }
  return jsonResponse({ admins: out });
}

function addAdmin(newName, addedByName) {
  if (!newName) return jsonResponse({ error: 'missing name' });
  const sheet = adminsSheet();
  const newKey = generateAdminKey();
  sheet.appendRow([newName, newKey, 'Admin', 'Yes', addedByName, new Date()]);
  return jsonResponse({ result: 'success', name: newName, key: newKey });
}

function setAdminActive(targetKey, active) {
  const admin = findAdminByKey(targetKey);
  if (!admin) return jsonResponse({ error: 'admin not found' });
  if (admin.role === 'Owner') return jsonResponse({ error: 'cannot deactivate an Owner' });
  const sheet = adminsSheet();
  sheet.getRange(admin.row, ACOL.ACTIVE).setValue(active ? 'Yes' : 'No');
  return jsonResponse({ result: 'success' });
}

function generateAdminKey() {
  return 'SIH-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

// ---------- Complaints: status lookup + list + update ----------

function getStatus(ticketId, teamNo) {
  if (!ticketId || !teamNo) return jsonResponse({ error: 'missing ticketId or teamNo' });
  const sheet = complaintsSheet();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (
      String(row[COL.TICKET - 1]).trim().toUpperCase() === String(ticketId).trim().toUpperCase() &&
      String(row[COL.TEAM_NO - 1]).trim().toUpperCase() === String(teamNo).trim().toUpperCase()
    ) {
      return jsonResponse({
        found: true,
        ticketId: row[COL.TICKET - 1],
        teamName: row[COL.TEAM_NAME - 1],
        category: row[COL.CATEGORY - 1],
        subject: row[COL.SUBJECT - 1],
        venue: row[COL.VENUE - 1],
        status: row[COL.STATUS - 1],
        remarks: row[COL.REMARKS - 1],
        lastUpdated: row[COL.UPDATED - 1]
      });
    }
  }
  return jsonResponse({ found: false });
}

function listComplaints() {
  const sheet = complaintsSheet();
  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[COL.TICKET - 1]) continue;
    out.push({
      ticketId: row[COL.TICKET - 1],
      timestamp: row[COL.TIMESTAMP - 1],
      teamName: row[COL.TEAM_NAME - 1],
      teamNo: row[COL.TEAM_NO - 1],
      category: row[COL.CATEGORY - 1],
      subject: row[COL.SUBJECT - 1],
      issue: row[COL.ISSUE - 1],
      venue: row[COL.VENUE - 1],
      status: row[COL.STATUS - 1],
      remarks: row[COL.REMARKS - 1],
      lastUpdated: row[COL.UPDATED - 1],
      handledBy: row[COL.HANDLED_BY - 1]
    });
  }
  out.reverse();
  return jsonResponse({ complaints: out });
}

function updateComplaint(params, adminName) {
  const ticketId = params.ticketId;
  if (!ticketId) return jsonResponse({ error: 'missing ticketId' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = complaintsSheet();
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][COL.TICKET - 1]).trim().toUpperCase() === String(ticketId).trim().toUpperCase()) {
        const rowIndex = i + 1;
        if (params.status) sheet.getRange(rowIndex, COL.STATUS).setValue(params.status);
        if (params.remarks !== undefined) sheet.getRange(rowIndex, COL.REMARKS).setValue(params.remarks);
        sheet.getRange(rowIndex, COL.UPDATED).setValue(new Date());
        sheet.getRange(rowIndex, COL.HANDLED_BY).setValue(adminName);
        return jsonResponse({ result: 'success' });
      }
    }
    return jsonResponse({ error: 'ticket not found' });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
