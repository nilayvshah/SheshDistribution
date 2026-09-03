// ════════════════════════════════════════════════════════
//  SZVS Jain Samaj — Notebook Distribution Entry System
//  Google Apps Script Backend
//
//  CHANGES FROM PREVIOUS VERSION
//  1. Entry ID allocated SERVER-SIDE inside a LockService
//     block. The browser no longer predicts IDs, so two
//     desks on the same login can no longer collide.
//  2. Duplicate check is LIFETIME (whole sheet), not
//     same-day, and is re-enforced server-side on insert.
//  3. Qty / Amount / Marksheet removed entirely.
//  4. OTP added to master_data, shown on screen and
//     stored on the transaction row.
//  5. print_required flag added to user_master.
// ════════════════════════════════════════════════════════

const MEMBERS_SHEET = 'master_data';
const ENTRIES_SHEET = 'transaction_data';
const USERS_SHEET   = 'user_master';

// ── master_data columns (0-based) ──
// A=0 memberid | B=1 new_code | C=2 name | D=3 address
// E=4 pincode  | F=5 mobile   | G=6 manav_rahat | H=7 otp

// ── user_master columns (0-based) ──
// A=0 userid | B=1 password | C=2 name | D=3 range_start
// E=4 range_end | F=5 active | G=6 is_admin | H=7 print_required
// I=8 dup_access

// ── transaction_data columns (1-based, compact layout) ──
// A=1 entry_id    | B=2 member_id | C=3 member_name
// D=4 otp         | E=5 entry_method (qrcode/manual)
// F=6 userid      | G=7 created_on
// H=8 remark      ('Duplicate Slip' or blank — system set)
// I=9 comment     (operator free text)
// J=10 changedby  | K=11 updatedon
const ENTRY_COLS = 11;

const ENTRY_HEADERS = [
  'Entry id', 'Member_id', 'Member Name', 'OTP', 'Entry Method',
  'Userid', 'Created on ( Date & time )', 'remark', 'comment',
  'changedby', 'updatedon'
];

const DUP_REMARK = 'Duplicate Slip';

// ════════════════════════════════════════════════════════
//  ENTRY POINT
// ════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter, action = p.action || '';
  try {
    switch (action) {
      case 'login':        return jsonResponse(loginUser(p.userid, p.password));
      case 'getMembers':   return jsonResponse(getMembers());
      case 'checkDup':     return jsonResponse(checkDuplicate(p.member_id));
      case 'addEntry':     return jsonResponse(addEntry(p));
      case 'getEntries':   return jsonResponse(getEntries(p.logged_by));
      case 'getAllEntries':return jsonResponse(getAllEntries(p.logged_by));
      case 'getSummary':   return jsonResponse(getSummary());
      case 'getUserList':  return jsonResponse(getUserList());
      case 'updateEntry':  return jsonResponse(updateEntry(p));
      default:             return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ════════════════════════════════════════════════════════
//  loginUser
//  No next_entry_id is returned any more — the server
//  allocates the ID at insert time instead.
// ════════════════════════════════════════════════════════
function loginUser(userid, password) {
  if (!userid || !password)
    return { success: false, error: 'Enter your username and password' };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) return { success: false, error: 'Sheet "' + USERS_SHEET + '" not found' };

  const rows = userSheet.getDataRange().getValues();
  let userRow = null;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).trim().toLowerCase() === userid.trim().toLowerCase()) {
      if (String(r[5]).trim().toUpperCase() === 'NO')
        return { success: false, error: 'This account is disabled. Contact the administrator.' };
      if (String(r[1]).trim() !== password.trim())
        return { success: false, error: 'Incorrect password' };
      userRow = r; break;
    }
  }
  if (!userRow) return { success: false, error: 'User not found' };

  // print_required: blank or YES → show print dialog, NO → skip it
  const printFlag = String(userRow[7] || '').trim().toUpperCase();

  return {
    success:        true,
    userid:         String(userRow[0]).trim(),
    name:           String(userRow[2]).trim(),
    range_start:    Number(userRow[3]),
    range_end:      Number(userRow[4]),
    is_admin:       String(userRow[6] || '').trim().toUpperCase() === 'YES',
    print_required: printFlag !== 'NO',
    dup_access:     hasDupAccess(userRow),
  };
}

// ════════════════════════════════════════════════════════
//  Duplicate slip access
//  Column I of user_master. Explicit grant: only YES
//  allows it, blank or anything else denies. Admin status
//  does not override this — set it per user.
// ════════════════════════════════════════════════════════
function hasDupAccess(userRow) {
  return String(userRow[8] || '').trim().toUpperCase() === 'YES';
}

function findUserRow(userid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  const target = String(userid).trim().toLowerCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === target) return rows[i];
  }
  return null;
}

// ════════════════════════════════════════════════════════
//  MEMBERS
// ════════════════════════════════════════════════════════
function loadAllMembers(ss) {
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];

  const data = sheet.getRange(2, 1, last - 1, 8).getValues();
  const out  = [];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim().toUpperCase();
    if (!id) continue;
    out.push({
      memberid:    id,
      new_code:    cleanNum(data[i][1]),
      name:        String(data[i][2]).trim(),
      address:     String(data[i][3]).trim(),
      pincode:     cleanNum(data[i][4]),
      mobile:      cleanNum(data[i][5]),
      manav_rahat: String(data[i][6] || '').trim(),
      otp:         cleanNum(data[i][7]),
    });
  }
  return out;
}

function getMembers() {
  const members = loadAllMembers(SpreadsheetApp.getActiveSpreadsheet());
  return { success: true, members: members, count: members.length };
}

// Sheet cells sometimes hold numbers where text was meant
// (pincode 380054.0, mobile 9.824082988E9). Render them as
// plain digits rather than scientific notation / decimals.
function cleanNum(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') {
    return Number.isInteger(v) ? String(v) : String(v).replace(/\.0+$/, '');
  }
  return String(v).trim();
}

// ════════════════════════════════════════════════════════
//  checkDuplicate — LIFETIME, across all users
//  Single getRange read instead of three.
// ════════════════════════════════════════════════════════
function checkDuplicate(memberId) {
  if (!memberId) return { duplicate: false };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { duplicate: false };

  const rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, ENTRY_COLS).getValues();
  const query = String(memberId).trim().toUpperCase();

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toUpperCase() !== query) continue;
    const ts = rows[i][6] ? new Date(rows[i][6]) : null;
    return {
      duplicate: true,
      entry_id:  rows[i][0],
      by:        String(rows[i][5]).trim(),
      remark:    String(rows[i][7] || '').trim(),
      at:        ts ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd MMM, hh:mm a') : '',
    };
  }
  return { duplicate: false };
}

// ════════════════════════════════════════════════════════
//  addEntry
//  Everything that must be atomic happens inside the lock:
//  duplicate re-check → ID allocation → append.
// ════════════════════════════════════════════════════════
function addEntry(p) {
  const req = ['memberid', 'member_name', 'entry_method', 'logged_by'];
  for (const f of req)
    if (!p[f] || String(p[f]).trim() === '')
      return { success: false, error: 'Missing: ' + f };

  const isDupSlip = String(p.remark || '').trim() === DUP_REMARK;
  const memberId  = String(p.memberid).trim().toUpperCase();

  // Duplicate slips need explicit permission. Checked here
  // and not only in the browser, since the endpoint can be
  // called directly.
  if (isDupSlip) {
    const userRow = findUserRow(p.logged_by);
    if (!userRow || !hasDupAccess(userRow))
      return { success: false, error: 'You do not have access to issue duplicate slips' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (e) {
    return { success: false, error: 'Server busy, try again' };
  }

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ENTRIES_SHEET);
    if (!sheet) return { success: false, error: 'Sheet "' + ENTRIES_SHEET + '" not found' };

    if (sheet.getLastRow() === 0) sheet.appendRow(ENTRY_HEADERS);

    const lastRow = sheet.getLastRow();
    const rows    = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, ENTRY_COLS).getValues()
      : [];

    // ── Duplicate re-check (skipped for duplicate slips) ──
    if (!isDupSlip) {
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim().toUpperCase() === memberId) {
          const ts = rows[i][6] ? new Date(rows[i][6]) : null;
          return {
            success:   false,
            duplicate: true,
            error:     'Already entered by ' + String(rows[i][5]).trim() +
                       (ts ? ' on ' + Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd MMM, hh:mm a') : ''),
          };
        }
      }
    }

    // ── Allocate the next ID inside this user's range ──
    const rangeStart = Number(p.range_start);
    const rangeEnd   = Number(p.range_end);
    if (!rangeStart || !rangeEnd)
      return { success: false, error: 'Entry ID range missing for this user' };

    let maxId = rangeStart - 1;
    for (let i = 0; i < rows.length; i++) {
      const id = Number(rows[i][0]);
      if (id >= rangeStart && id <= rangeEnd && id > maxId) maxId = id;
    }
    const entryId = maxId + 1;
    if (entryId > rangeEnd)
      return { success: false, error: 'Entry ID range exhausted for this user' };

    const now = new Date();
    sheet.appendRow([
      entryId,                              // A entry_id
      memberId,                             // B member_id
      String(p.member_name).trim(),         // C member_name
      String(p.otp || '').trim(),           // D otp
      p.entry_method,                       // E entry_method
      p.logged_by,                          // F userid
      now,                                  // G created_on
      isDupSlip ? DUP_REMARK : '',          // H remark
      String(p.comment || '').trim(),       // I comment
      '',                                   // J changedby
      '',                                   // K updatedon
    ]);

    return { success: true, entry_id: entryId, timestamp: now.toISOString() };

  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
//  REPORTS
// ════════════════════════════════════════════════════════
function getEntries(loggedBy) {
  if (!loggedBy) return { entries: [] };
  const target = loggedBy.toLowerCase();
  return { entries: readEntrySheet().filter(r => r.userid.toLowerCase() === target) };
}

function getAllEntries(loggedBy) {
  const rows = readEntrySheet();
  if (!loggedBy || loggedBy.toUpperCase() === 'ALL') return { entries: rows };
  return { entries: rows.filter(r => r.userid.toLowerCase() === loggedBy.toLowerCase()) };
}

// User-wise counts. No amounts any more — this is now
// entries / duplicate slips / total.
function getSummary() {
  const rows = readEntrySheet();

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  const nameMap   = {};
  if (userSheet) {
    const ud = userSheet.getDataRange().getValues();
    for (let i = 1; i < ud.length; i++)
      nameMap[String(ud[i][0]).trim().toLowerCase()] = String(ud[i][2]).trim();
  }

  const agg = {};
  for (const r of rows) {
    const uid = r.userid.toLowerCase();
    if (!agg[uid]) {
      agg[uid] = {
        userid: r.userid, name: nameMap[uid] || r.userid,
        count: 0, dup_count: 0, normal_count: 0,
      };
    }
    agg[uid].count++;
    if (r.remark === DUP_REMARK) agg[uid].dup_count++;
    else                         agg[uid].normal_count++;
  }

  const summary = Object.values(agg);
  const grand = {
    userid: '__TOTAL__', name: 'GRAND TOTAL',
    count:        summary.reduce((s, r) => s + r.count, 0),
    dup_count:    summary.reduce((s, r) => s + r.dup_count, 0),
    normal_count: summary.reduce((s, r) => s + r.normal_count, 0),
  };

  return { summary, grand };
}

function getUserList() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) return { users: [] };

  const rows  = userSheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).trim().toUpperCase() === 'NO') continue;
    users.push({ userid: String(rows[i][0]).trim(), name: String(rows[i][2]).trim() });
  }
  return { users };
}

// ════════════════════════════════════════════════════════
//  updateEntry — admin only
//  Member ID, name and OTP are never changed. Editable:
//  entry_method, userid, remark, comment.
// ════════════════════════════════════════════════════════
function updateEntry(p) {
  if (!p.entry_id) return { success: false, error: 'entry_id required' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet) return { success: false, error: 'Entries sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No entries found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  // Match on entry_id AND member_id. Legacy data contains
  // repeated entry_ids; this stops an edit landing on the
  // wrong row when that happens.
  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== String(p.entry_id).trim()) continue;
    if (p.member_id && String(rows[i][1]).trim().toUpperCase() !== String(p.member_id).trim().toUpperCase()) continue;
    targetRow = i + 2;
    break;
  }
  if (targetRow === -1)
    return { success: false, error: 'Entry ' + p.entry_id + ' not found' };

  const remark = String(p.remark || '').trim() === DUP_REMARK ? DUP_REMARK : '';

  sheet.getRange(targetRow, 5).setValue(p.entry_method || 'manual');  // E
  sheet.getRange(targetRow, 6).setValue(p.userid  || '');             // F
  sheet.getRange(targetRow, 8).setValue(remark);                      // H
  sheet.getRange(targetRow, 9).setValue(p.comment || '');             // I
  sheet.getRange(targetRow, 10).setValue(p.changedby || '');          // J
  sheet.getRange(targetRow, 11).setValue(new Date());                 // K

  return { success: true, entry_id: p.entry_id, remark: remark };
}

// ════════════════════════════════════════════════════════
//  SHARED
// ════════════════════════════════════════════════════════
function readEntrySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, ENTRY_COLS).getValues();
  const entries = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    entries.push({
      entry_id:     r[0],
      member_id:    String(r[1]).trim(),
      member_name:  String(r[2]).trim(),
      otp:          cleanNum(r[3]),
      entry_method: String(r[4] || '').trim(),
      userid:       String(r[5]).trim(),
      created_on:   r[6] ? new Date(r[6]).toISOString() : '',
      remark:       String(r[7] || '').trim(),
      comment:      String(r[8] || '').trim(),
      changedby:    String(r[9] || '').trim(),
      changed_on:   r[10] ? new Date(r[10]).toISOString() : '',
    });
  }
  return entries;
}

function jsonResponse(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
