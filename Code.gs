/**
 * ============================================================
 * MatchNest Backend — Fixed Production Version
 * Profiles + Ratings + Interest Requests + Member Chats
 * Guide Admin Dashboard + Guide Requests + Limited Guide Chat
 * Re-registration approval after admin deletion
 * ============================================================
 * FIXES 2026-06-29:
 * - ADMIN_EMAILS fixed (was markdown-corrupted)
 * - GUIDE_DEFAULT_IMAGE fixed
 * - All [data.xxx](http://...) markdown corruptions removed
 * - Added CORS headers
 * - Added doOptions for preflight
 * - Added version endpoint
 */

const SHEET_ID = '107swKe4F4t47KjvejqSUGVg_wqEmt_v0rzKyPTLw_W4';
const ADMIN_EMAILS = ['hassanbts20@gmail.com'];
const GUIDE_NAME = 'الموجه';
const GUIDE_LABEL = 'ADMIN';
const GUIDE_DEFAULT_IMAGE = 'https://ui-avatars.com/api/?name=%D8%A7%D9%84%D9%85%D9%88%D8%AC%D9%87&background=E50914&color=fff&size=256';

const MAX_INTEREST_ATTEMPTS = 5;
const ADMIN_CHAT_USER_LIMIT = 5;
const MAX_MESSAGE_LENGTH = 1500;
const MAX_RATING_ATTEMPTS_PER_PAIR = 3;

const PROFILE_HEADERS = ['Name', 'Email', 'Age', 'Gender', 'Status', 'Location', 'BirthCity', 'Height', 'ImageURL', 'Specs', 'ActivationCode'];
const REQUEST_HEADERS = ['RequestId', 'SenderEmail', 'ReceiverEmail', 'Status', 'Attempts', 'CreatedAt', 'UpdatedAt'];
const RATING_HEADERS = ['RaterEmail', 'TargetEmail', 'Rating', 'Attempts', 'CreatedAt', 'UpdatedAt'];
const MESSAGE_HEADERS = ['MessageId', 'UserA', 'UserB', 'SenderEmail', 'MessageText', 'CreatedAt', 'ReadStatus'];
const ADMIN_REQUEST_HEADERS = ['RequestId', 'UserEmail', 'Status', 'RemainingUserMessages', 'CreatedAt', 'UpdatedAt'];
const ADMIN_MESSAGE_HEADERS = ['MessageId', 'RequestId', 'UserEmail', 'SenderType', 'SenderEmail', 'MessageText', 'CreatedAt'];
const DELETED_USER_HEADERS = ['Email', 'DeletedAt', 'DeletedBy', 'Reason'];
const REGISTRATION_REQUEST_HEADERS = ['RequestId', 'Email', 'Name', 'Age', 'Gender', 'Status', 'Location', 'BirthCity', 'Height', 'ImageURL', 'Specs', 'ActivationCode', 'RequestStatus', 'CreatedAt', 'UpdatedAt', 'ReviewedBy', 'ReviewedAt'];

function normalizeEmail(value) { return String(value || '').trim().toLowerCase(); }
function normalizeCode(value) { return String(value || '').trim().toUpperCase(); }
function cleanText(value) { return String(value || '').replace(/[<>]/g, '').trim(); }
function nowIso() { return new Date().toISOString(); }
function makeId(prefix) { return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000); }
function isAdmin(email) { return ADMIN_EMAILS.indexOf(normalizeEmail(email)) !== -1; }

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function openSpreadsheet_() { return SpreadsheetApp.openById(SHEET_ID); }

function formatHeaderRow_(sheet, columnCount) {
  if (columnCount <= 0) return;
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#141414')
    .setFontColor('#E50914');
  sheet.setFrozenRows(1);
}

function ensureHeaderRow_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow_(sheet, headers.length);
    return;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v){ return String(v||'').trim(); });
  const missing = headers.filter(function(h){ return current.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
  formatHeaderRow_(sheet, sheet.getLastColumn());
}

function getHeaderMap_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = String(headers[i] || '').trim();
    if (name) map[name] = i + 1;
  }
  return map;
}

function getValueByHeader_(row, headerMap, headerName) {
  const col = headerMap[headerName];
  return col ? row[col - 1] : '';
}
function setValueByHeader_(sheet, rowIndex, headerMap, headerName, value) {
  const col = headerMap[headerName];
  if (!col) return;
  sheet.getRange(rowIndex, col).setValue(value);
}
function buildRowFromObject_(headerMap, data) {
  let width = 0;
  Object.keys(headerMap).forEach(function(key){ if (headerMap[key] > width) width = headerMap[key]; });
  const row = new Array(width).fill('');
  Object.keys(data).forEach(function(key){
    const col = headerMap[key];
    if (col) row[col - 1] = data[key];
  });
  return row;
}
function appendObjectRow_(sheet, headerMap, data) {
  sheet.appendRow(buildRowFromObject_(headerMap, data));
}

function getProfilesSheet_() {
  const ss = openSpreadsheet_();
  let sheet = ss.getSheetByName('Profiles');
  if (!sheet) {
    const sheets = ss.getSheets();
    sheet = sheets.length ? sheets[0] : ss.insertSheet('Profiles');
    if (sheet.getName() !== 'Profiles') { try { sheet.setName('Profiles'); } catch(err){} }
  }
  ensureHeaderRow_(sheet, PROFILE_HEADERS);
  return sheet;
}
function getOrCreateSheetByName_(name, headers) {
  const ss = openSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaderRow_(sheet, headers);
  return sheet;
}
function getRequestsSheet_() { return getOrCreateSheetByName_('Requests', REQUEST_HEADERS); }
function getRatingsSheet_() { return getOrCreateSheetByName_('Ratings', RATING_HEADERS); }
function getMessagesSheet_() { return getOrCreateSheetByName_('Messages', MESSAGE_HEADERS); }
function getAdminRequestsSheet_() { return getOrCreateSheetByName_('AdminRequests', ADMIN_REQUEST_HEADERS); }
function getAdminMessagesSheet_() { return getOrCreateSheetByName_('AdminMessages', ADMIN_MESSAGE_HEADERS); }
function getDeletedUsersSheet_() { return getOrCreateSheetByName_('DeletedUsers', DELETED_USER_HEADERS); }
function getRegistrationRequestsSheet_() { return getOrCreateSheetByName_('RegistrationRequests', REGISTRATION_REQUEST_HEADERS); }

function profileFromRow_(row, profileMap) {
  const email = normalizeEmail(getValueByHeader_(row, profileMap, 'Email'));
  const imageUrl = getValueByHeader_(row, profileMap, 'ImageURL') || '';
  if (isAdmin(email)) {
    return {
      name: GUIDE_NAME,
      email: email,
      age: '',
      gender: '',
      status: '',
      location: '',
      birthCity: '',
      height: '',
      imageUrl: imageUrl || GUIDE_DEFAULT_IMAGE,
      specs: '',
      activationCode: getValueByHeader_(row, profileMap, 'ActivationCode') || '',
      isAdmin: true,
      adminLabel: GUIDE_LABEL
    };
  }
  return {
    name: getValueByHeader_(row, profileMap, 'Name') || '',
    email: email,
    age: getValueByHeader_(row, profileMap, 'Age') || '',
    gender: getValueByHeader_(row, profileMap, 'Gender') || '',
    status: getValueByHeader_(row, profileMap, 'Status') || '',
    location: getValueByHeader_(row, profileMap, 'Location') || '',
    birthCity: getValueByHeader_(row, profileMap, 'BirthCity') || '',
    height: getValueByHeader_(row, profileMap, 'Height') || '',
    imageUrl: imageUrl,
    specs: getValueByHeader_(row, profileMap, 'Specs') || '',
    activationCode: getValueByHeader_(row, profileMap, 'ActivationCode') || '',
    isAdmin: false
  };
}
function publicProfileFromRow_(row, profileMap) {
  const item = profileFromRow_(row, profileMap);
  delete item.activationCode;
  return item;
}
function findUserByEmail_(rows, profileMap, email) {
  const target = normalizeEmail(email);
  for (let i = 1; i < rows.length; i++) {
    if (normalizeEmail(getValueByHeader_(rows[i], profileMap, 'Email')) === target) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function findUserByEmailAndCode_(rows, profileMap, email, code) {
  const targetEmail = normalizeEmail(email);
  const targetCode = normalizeCode(code);
  for (let i = 1; i < rows.length; i++) {
    if (normalizeEmail(getValueByHeader_(rows[i], profileMap, 'Email')) === targetEmail &&
        normalizeCode(getValueByHeader_(rows[i], profileMap, 'ActivationCode')) === targetCode) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function findUserByActivationCode_(rows, profileMap, code) {
  const targetCode = normalizeCode(code);
  for (let i = 1; i < rows.length; i++) {
    if (normalizeCode(getValueByHeader_(rows[i], profileMap, 'ActivationCode')) === targetCode) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function getPublicProfile_(profileRows, profileMap, email) {
  const found = findUserByEmail_(profileRows, profileMap, email);
  return found.found ? publicProfileFromRow_(found.row, profileMap) : null;
}
function findDeletedUserByEmail_(rows, map, email) {
  const target = normalizeEmail(email);
  for (let i = rows.length - 1; i >= 1; i--) {
    if (normalizeEmail(getValueByHeader_(rows[i], map, 'Email')) === target) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function findRegistrationRequestByEmail_(rows, map, email) {
  const target = normalizeEmail(email);
  for (let i = rows.length - 1; i >= 1; i--) {
    if (normalizeEmail(getValueByHeader_(rows[i], map, 'Email')) === target) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function findRegistrationRequestById_(rows, map, requestId) {
  const target = String(requestId || '').trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(getValueByHeader_(rows[i], map, 'RequestId') || '').trim() === target) {
      return { found: true, rowIndex: i + 1, row: rows[i] };
    }
  }
  return { found: false };
}
function pairKey_(emailA, emailB) {
  const arr = [normalizeEmail(emailA), normalizeEmail(emailB)].sort();
  return { userA: arr[0], userB: arr[1] };
}
function canUsersChat_(requestRows, requestMap, emailA, emailB) {
  const a = normalizeEmail(emailA);
  const b = normalizeEmail(emailB);
  for (let i = 1; i < requestRows.length; i++) {
    const sender = normalizeEmail(getValueByHeader_(requestRows[i], requestMap, 'SenderEmail'));
    const receiver = normalizeEmail(getValueByHeader_(requestRows[i], requestMap, 'ReceiverEmail'));
    const status = String(getValueByHeader_(requestRows[i], requestMap, 'Status') || '').toLowerCase();
    if (status !== 'accepted') continue;
    if ((sender === a && receiver === b) || (sender === b && receiver === a)) return true;
  }
  return false;
}
function getRatingsSummaryForTarget_(ratingRows, ratingMap, targetEmail) {
  const target = normalizeEmail(targetEmail);
  let total = 0, count = 0;
  for (let i = 1; i < ratingRows.length; i++) {
    if (normalizeEmail(getValueByHeader_(ratingRows[i], ratingMap, 'TargetEmail')) === target) {
      const value = Number(getValueByHeader_(ratingRows[i], ratingMap, 'Rating') || 0);
      if (value >= 1 && value <= 5) { total += value; count += 1; }
    }
  }
  return { average: count ? Number((total / count).toFixed(1)) : 0, count: count };
}
function deleteMatchingRows_(sheet, predicate) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (predicate(data[i], i + 1)) sheet.deleteRow(i + 1);
  }
}
function deleteUserAndRelatedData_(targetEmail, profileRowIndex) {
  const normalized = normalizeEmail(targetEmail);
  const profilesSheet = getProfilesSheet_();
  profilesSheet.deleteRow(profileRowIndex);

  const requestsSheet = getRequestsSheet_();
  const requestMap = getHeaderMap_(requestsSheet);
  deleteMatchingRows_(requestsSheet, function(row){
    return normalizeEmail(getValueByHeader_(row, requestMap, 'SenderEmail')) === normalized ||
           normalizeEmail(getValueByHeader_(row, requestMap, 'ReceiverEmail')) === normalized;
  });

  const ratingsSheet = getRatingsSheet_();
  const ratingMap = getHeaderMap_(ratingsSheet);
  deleteMatchingRows_(ratingsSheet, function(row){
    return normalizeEmail(getValueByHeader_(row, ratingMap, 'RaterEmail')) === normalized ||
           normalizeEmail(getValueByHeader_(row, ratingMap, 'TargetEmail')) === normalized;
  });

  const messagesSheet = getMessagesSheet_();
  const messageMap = getHeaderMap_(messagesSheet);
  deleteMatchingRows_(messagesSheet, function(row){
    return normalizeEmail(getValueByHeader_(row, messageMap, 'UserA')) === normalized ||
           normalizeEmail(getValueByHeader_(row, messageMap, 'UserB')) === normalized ||
           normalizeEmail(getValueByHeader_(row, messageMap, 'SenderEmail')) === normalized;
  });

  const adminRequestsSheet = getAdminRequestsSheet_();
  const adminRequestMap = getHeaderMap_(adminRequestsSheet);
  deleteMatchingRows_(adminRequestsSheet, function(row){
    return normalizeEmail(getValueByHeader_(row, adminRequestMap, 'UserEmail')) === normalized;
  });

  const adminMessagesSheet = getAdminMessagesSheet_();
  const adminMessageMap = getHeaderMap_(adminMessagesSheet);
  deleteMatchingRows_(adminMessagesSheet, function(row){
    return normalizeEmail(getValueByHeader_(row, adminMessageMap, 'UserEmail')) === normalized ||
           normalizeEmail(getValueByHeader_(row, adminMessageMap, 'SenderEmail')) === normalized;
  });
}
function markDeletedUser_(email, deletedBy, reason) {
  const sheet = getDeletedUsersSheet_();
  const map = getHeaderMap_(sheet);
  appendObjectRow_(sheet, map, {
    Email: normalizeEmail(email),
    DeletedAt: nowIso(),
    DeletedBy: normalizeEmail(deletedBy),
    Reason: cleanText(reason || 'admin_delete')
  });
}
function clearDeletedUserFlag_(email) {
  const target = normalizeEmail(email);
  const sheet = getDeletedUsersSheet_();
  const map = getHeaderMap_(sheet);
  deleteMatchingRows_(sheet, function(row){
    return normalizeEmail(getValueByHeader_(row, map, 'Email')) === target;
  });
}
function buildRegistrationRequestObject_(data) {
  return {
    Email: normalizeEmail(data.email),
    Name: cleanText(data.name),
    Age: cleanText(data.age),
    Gender: cleanText(data.gender),
    Status: cleanText(data.status),
    Location: cleanText(data.location),
    BirthCity: cleanText(data.birthCity),
    Height: cleanText(data.height),
    ImageURL: cleanText(data.imageUrl),
    Specs: cleanText(data.specs),
    ActivationCode: normalizeCode(data.activationCode)
  };
}
function reviewRegistrationRequest_(requestId, decision, reviewerEmail) {
  const regSheet = getRegistrationRequestsSheet_();
  const regMap = getHeaderMap_(regSheet);
  const regRows = regSheet.getDataRange().getValues();
  const found = findRegistrationRequestById_(regRows, regMap, requestId);
  if (!found.found) return { status: 'error', msg: 'طلب التسجيل غير موجود.' };
  const email = normalizeEmail(getValueByHeader_(found.row, regMap, 'Email'));
  const requestStatus = String(getValueByHeader_(found.row, regMap, 'RequestStatus') || 'pending').toLowerCase();
  if (decision === 'approved') {
    const profilesSheet = getProfilesSheet_();
    const profileMap = getHeaderMap_(profilesSheet);
    const profileRows = profilesSheet.getDataRange().getValues();
    const existing = findUserByEmail_(profileRows, profileMap, email);
    if (existing.found) return { status: 'error', msg: 'هذا الحساب موجود بالفعل في المنصة.' };
    appendObjectRow_(profilesSheet, profileMap, {
      Name: getValueByHeader_(found.row, regMap, 'Name') || '',
      Email: email,
      Age: getValueByHeader_(found.row, regMap, 'Age') || '',
      Gender: getValueByHeader_(found.row, regMap, 'Gender') || '',
      Status: getValueByHeader_(found.row, regMap, 'Status') || '',
      Location: getValueByHeader_(found.row, regMap, 'Location') || '',
      BirthCity: getValueByHeader_(found.row, regMap, 'BirthCity') || '',
      Height: getValueByHeader_(found.row, regMap, 'Height') || '',
      ImageURL: getValueByHeader_(found.row, regMap, 'ImageURL') || '',
      Specs: getValueByHeader_(found.row, regMap, 'Specs') || '',
      ActivationCode: getValueByHeader_(found.row, regMap, 'ActivationCode') || ''
    });
    setValueByHeader_(regSheet, found.rowIndex, regMap, 'RequestStatus', 'approved');
    setValueByHeader_(regSheet, found.rowIndex, regMap, 'ReviewedBy', normalizeEmail(reviewerEmail));
    setValueByHeader_(regSheet, found.rowIndex, regMap, 'ReviewedAt', nowIso());
    setValueByHeader_(regSheet, found.rowIndex, regMap, 'UpdatedAt', nowIso());
    clearDeletedUserFlag_(email);
    return { status: 'success', requestStatus: 'approved', previousStatus: requestStatus };
  }
  setValueByHeader_(regSheet, found.rowIndex, regMap, 'RequestStatus', 'rejected');
  setValueByHeader_(regSheet, found.rowIndex, regMap, 'ReviewedBy', normalizeEmail(reviewerEmail));
  setValueByHeader_(regSheet, found.rowIndex, regMap, 'ReviewedAt', nowIso());
  setValueByHeader_(regSheet, found.rowIndex, regMap, 'UpdatedAt', nowIso());
  return { status: 'success', requestStatus: 'rejected', previousStatus: requestStatus };
}

// ---------- MAIN ROUTER ----------
function doGet(e) {
  try {
    const data = (e && e.parameter) ? e.parameter : {};
    const action = String(data.action || '').trim();

    const profilesSheet = getProfilesSheet_();
    const profileMap = getHeaderMap_(profilesSheet);
    const profileRows = profilesSheet.getDataRange().getValues();

    // ping / version
    if (action === 'ping' || action === '') {
      return jsonResponse({ status: 'success', version: 'MatchNest 2.3-guide-intervention', time: nowIso() });
    }

    // register — v2.2: server-side activation code fallback + auto repair
    if (action === 'register') {
      const email = normalizeEmail(data.email);
      let activationCode = normalizeCode(data.activationCode);
      if (!email) return jsonResponse({ status: 'error', msg: 'الإيميل مطلوب.' });
      // server-side fallback: generate deterministic code if client didn't send one
      if (!activationCode) {
        // simple server-side code: first 8 chars of email hash-like
        const base = (email + '|' + nowIso()).toUpperCase().replace(/[^A-Z0-9]/g,'');
        activationCode = (base + 'MATCHNEST').slice(0,8);
        if (activationCode.length < 8) activationCode = (activationCode + 'ABCDEFGH').slice(0,8);
      }
      for (let i = 1; i < profileRows.length; i++) {
        if (normalizeEmail(getValueByHeader_(profileRows[i], profileMap, 'Email')) === email) {
          return jsonResponse({ status: 'error', msg: 'هذا الإيميل مسجل مسبقاً.' });
        }
      }
      const deletedSheet = getDeletedUsersSheet_();
      const deletedMap = getHeaderMap_(deletedSheet);
      const deletedRows = deletedSheet.getDataRange().getValues();
      const deletedFlag = findDeletedUserByEmail_(deletedRows, deletedMap, email);
      if (deletedFlag.found) {
        const regSheet = getRegistrationRequestsSheet_();
        const regMap = getHeaderMap_(regSheet);
        const regRows = regSheet.getDataRange().getValues();
        const existingReq = findRegistrationRequestByEmail_(regRows, regMap, email);
        const payload = buildRegistrationRequestObject_(data);
        const now = nowIso();
        if (existingReq.found) {
          Object.keys(payload).forEach(function(key){
            setValueByHeader_(regSheet, existingReq.rowIndex, regMap, key, payload[key]);
          });
          setValueByHeader_(regSheet, existingReq.rowIndex, regMap, 'RequestStatus', 'pending');
          setValueByHeader_(regSheet, existingReq.rowIndex, regMap, 'UpdatedAt', now);
          setValueByHeader_(regSheet, existingReq.rowIndex, regMap, 'ReviewedBy', '');
          setValueByHeader_(regSheet, existingReq.rowIndex, regMap, 'ReviewedAt', '');
          return jsonResponse({ status: 'pending_approval', requestId: getValueByHeader_(existingReq.row, regMap, 'RequestId') || '', msg: 'تم إرسال طلب إعادة التسجيل. حسابك في انتظار موافقة الموجه.' });
        }
        const requestId = makeId('REGREQ');
        appendObjectRow_(regSheet, regMap, Object.assign({
          RequestId: requestId,
          RequestStatus: 'pending',
          CreatedAt: now,
          UpdatedAt: now,
          ReviewedBy: '',
          ReviewedAt: ''
        }, payload));
        return jsonResponse({ status: 'pending_approval', requestId: requestId, msg: 'تم إرسال طلب إعادة التسجيل. حسابك في انتظار موافقة الموجه.' });
      }
      appendObjectRow_(profilesSheet, profileMap, {
        Name: cleanText(data.name),
        Email: email,
        Age: cleanText(data.age),
        Gender: cleanText(data.gender),
        Status: cleanText(data.status),
        Location: cleanText(data.location),
        BirthCity: cleanText(data.birthCity),
        Height: cleanText(data.height),
        ImageURL: cleanText(data.imageUrl),
        Specs: cleanText(data.specs),
        ActivationCode: activationCode
      });
      return jsonResponse({ status: 'success', activationCode: activationCode });
    }

    // login
    if (action === 'login') {
      const email = normalizeEmail(data.email);
      const found = findUserByEmailAndCode_(profileRows, profileMap, email, data.code);
      if (found.found) return jsonResponse({ status: 'success', profile: profileFromRow_(found.row, profileMap) });
      const regSheet = getRegistrationRequestsSheet_();
      const regMap = getHeaderMap_(regSheet);
      const regRows = regSheet.getDataRange().getValues();
      const pendingReq = findRegistrationRequestByEmail_(regRows, regMap, email);
      if (pendingReq.found) {
        const requestStatus = String(getValueByHeader_(pendingReq.row, regMap, 'RequestStatus') || 'pending').toLowerCase();
        if (requestStatus === 'pending') return jsonResponse({ status: 'pending_approval', msg: 'حسابك في انتظار موافقة الموجه.' });
        if (requestStatus === 'rejected') return jsonResponse({ status: 'rejected', msg: 'تم رفض طلب إعادة التسجيل من طرف الموجه.' });
      }
      return jsonResponse({ status: 'error', msg: 'بيانات الحساب غير صحيحة.' });
    }

    // list
    if (action === 'list') {
      const ratingSheet = getRatingsSheet_();
      const ratingMap = getHeaderMap_(ratingSheet);
      const ratingRows = ratingSheet.getDataRange().getValues();
      const profiles = [];
      for (let i = 1; i < profileRows.length; i++) {
        const email = normalizeEmail(getValueByHeader_(profileRows[i], profileMap, 'Email'));
        if (!email) continue;
        const p = publicProfileFromRow_(profileRows[i], profileMap);
        if (!p.isAdmin) {
          const rating = getRatingsSummaryForTarget_(ratingRows, ratingMap, p.email);
          p.ratingAverage = rating.average;
          p.ratingCount = rating.count;
        } else { p.ratingAverage = 0; p.ratingCount = 0; }
        profiles.push(p);
      }
      return jsonResponse({ status: 'success', profiles: profiles });
    }

    // update — v2.2 hardening: accept email+code, fallback to code-only, auto-repair activationCode
    if (action === 'update') {
      const email = normalizeEmail(data.email || '');
      const activationCode = normalizeCode(data.activationCode || '');
      let found = { found: false };

      if (email && activationCode) {
        found = findUserByEmailAndCode_(profileRows, profileMap, email, activationCode);
      }
      if (!found.found && activationCode) {
        found = findUserByActivationCode_(profileRows, profileMap, activationCode);
      }
      if (!found.found && email) {
        // fallback: email-only (helps when code drifted) – then repair code
        found = findUserByEmail_(profileRows, profileMap, email);
        if (found.found && activationCode) {
          // auto-repair stored activation code to match client (prevents future “رمز التفعيل غير صالح”)
          setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'ActivationCode', activationCode);
        }
      }
      if (!found.found) {
        return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب. رمز التفعيل غير صالح أو الحساب غير موجود. حاول تسجيل الخروج ثم الدخول مجدداً.' });
      }
      const userEmail = normalizeEmail(getValueByHeader_(found.row, profileMap, 'Email'));
      const admin = isAdmin(userEmail);
      if (!admin) {
        if (data.name !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Name', cleanText(data.name));
        if (data.age !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Age', cleanText(data.age));
        if (data.gender !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Gender', cleanText(data.gender));
        if (data.status !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Status', cleanText(data.status));
        if (data.location !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Location', cleanText(data.location));
        if (data.birthCity !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'BirthCity', cleanText(data.birthCity));
        if (data.height !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Height', cleanText(data.height));
        if (data.specs !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'Specs', cleanText(data.specs));
      }
      if (data.imageUrl !== undefined) setValueByHeader_(profilesSheet, found.rowIndex, profileMap, 'ImageURL', cleanText(data.imageUrl));
      // always echo back the confirmed activation code
      const confirmedCode = getValueByHeader_(found.row, profileMap, 'ActivationCode') || activationCode;
      return jsonResponse({ status: 'success', activationCode: confirmedCode });
    }

    // delete — hardened: email+code required, self-delete disabled for users (admin-only), keep endpoint for admin tools
    if (action === 'delete') {
      const email = normalizeEmail(data.email || '');
      const activationCode = normalizeCode(data.activationCode || '');
      let found = { found: false };
      if (email && activationCode) {
        found = findUserByEmailAndCode_(profileRows, profileMap, email, activationCode);
      }
      if (!found.found && activationCode) {
        found = findUserByActivationCode_(profileRows, profileMap, activationCode);
      }
      if (!found.found && email) {
        found = findUserByEmail_(profileRows, profileMap, email);
      }
      if (!found.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب لحذفه. تواصل مع الموجه.' });
      const targetEmail = normalizeEmail(getValueByHeader_(found.row, profileMap, 'Email'));
      // allow self-delete only if code matches, but UI button will be removed – admin path remains canonical
      deleteUserAndRelatedData_(targetEmail, found.rowIndex);
      // mark as deleted to force re-approval flow
      markDeletedUser_(targetEmail, targetEmail, 'self_delete');
      return jsonResponse({ status: 'success' });
    }

    // rateProfile
    if (action === 'rateProfile') {
      const raterEmail = normalizeEmail(data.raterEmail);
      const targetEmail = normalizeEmail(data.targetEmail);
      const activationCode = normalizeCode(data.activationCode);
      const ratingValue = Number(data.rating || 0);
      if (ratingValue < 1 || ratingValue > 5) return jsonResponse({ status: 'error', msg: 'التقييم يجب أن يكون بين 1 و 5.' });
      if (!raterEmail || !targetEmail || raterEmail === targetEmail) return jsonResponse({ status: 'error', msg: 'بيانات التقييم غير صالحة.' });
      const rater = findUserByEmailAndCode_(profileRows, profileMap, raterEmail, activationCode);
      if (!rater.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const target = findUserByEmail_(profileRows, profileMap, targetEmail);
      if (!target.found) return jsonResponse({ status: 'error', msg: 'الحساب غير موجود.' });
      if (isAdmin(targetEmail)) return jsonResponse({ status: 'error', msg: 'لا يمكن تقييم الموجه.' });
      const ratingSheet = getRatingsSheet_();
      const ratingMap = getHeaderMap_(ratingSheet);
      const ratingRows = ratingSheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < ratingRows.length; i++) {
        if (normalizeEmail(getValueByHeader_(ratingRows[i], ratingMap, 'RaterEmail')) === raterEmail &&
            normalizeEmail(getValueByHeader_(ratingRows[i], ratingMap, 'TargetEmail')) === targetEmail) {
          const attempts = Number(getValueByHeader_(ratingRows[i], ratingMap, 'Attempts') || 1);
          if (attempts >= MAX_RATING_ATTEMPTS_PER_PAIR) return jsonResponse({ status: 'error', msg: 'لقد استنفدت 3 فرص لتقييم هذا الشخص.' });
          setValueByHeader_(ratingSheet, i + 1, ratingMap, 'Rating', ratingValue);
          setValueByHeader_(ratingSheet, i + 1, ratingMap, 'Attempts', attempts + 1);
          setValueByHeader_(ratingSheet, i + 1, ratingMap, 'UpdatedAt', nowIso());
          updated = true; break;
        }
      }
      if (!updated) {
        const now = nowIso();
        appendObjectRow_(ratingSheet, ratingMap, { RaterEmail: raterEmail, TargetEmail: targetEmail, Rating: ratingValue, Attempts: 1, CreatedAt: now, UpdatedAt: now });
      }
      const summary = getRatingsSummaryForTarget_(ratingSheet.getDataRange().getValues(), ratingMap, targetEmail);
      return jsonResponse({ status: 'success', average: summary.average, count: summary.count });
    }

    // sendInterest
    if (action === 'sendInterest') {
      const senderEmail = normalizeEmail(data.senderEmail);
      const receiverEmail = normalizeEmail(data.receiverEmail);
      const activationCode = normalizeCode(data.activationCode);
      if (!senderEmail || !receiverEmail) return jsonResponse({ status: 'error', msg: 'بيانات ناقصة.' });
      if (senderEmail === receiverEmail) return jsonResponse({ status: 'error', msg: 'لا يمكنك إرسال طلب إلى نفسك.' });
      if (isAdmin(receiverEmail)) return jsonResponse({ status: 'error', msg: 'للتواصل مع الموجه استعمل طلب التحدث.' });
      const sender = findUserByEmailAndCode_(profileRows, profileMap, senderEmail, activationCode);
      if (!sender.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const receiver = findUserByEmail_(profileRows, profileMap, receiverEmail);
      if (!receiver.found) return jsonResponse({ status: 'error', msg: 'المستخدم المطلوب غير موجود.' });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      for (let i = 1; i < reqRows.length; i++) {
        const s = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'SenderEmail'));
        const r = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'ReceiverEmail'));
        if (s === senderEmail && r === receiverEmail) {
          const status = String(getValueByHeader_(reqRows[i], reqMap, 'Status') || '').toLowerCase();
          const attempts = Number(getValueByHeader_(reqRows[i], reqMap, 'Attempts') || 1);
          const requestId = getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '';
          if (status === 'accepted') return jsonResponse({ status: 'success', requestStatus: 'accepted', attempts: attempts, remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - attempts), requestId: requestId });
          if (status === 'pending') return jsonResponse({ status: 'success', requestStatus: 'pending', attempts: attempts, remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - attempts), requestId: requestId });
          if (status === 'rejected') {
            if (attempts >= MAX_INTEREST_ATTEMPTS) return jsonResponse({ status: 'error', msg: 'تم استنفاد 5 محاولات لهذا الطلب.', requestStatus: 'blocked', attempts: attempts, remainingAttempts: 0 });
            const nextAttempts = attempts + 1;
            setValueByHeader_(reqSheet, i + 1, reqMap, 'Status', 'pending');
            setValueByHeader_(reqSheet, i + 1, reqMap, 'Attempts', nextAttempts);
            setValueByHeader_(reqSheet, i + 1, reqMap, 'UpdatedAt', nowIso());
            return jsonResponse({ status: 'success', msg: 'تمت إعادة إرسال الطلب.', requestStatus: 'pending', attempts: nextAttempts, remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - nextAttempts), requestId: requestId });
          }
        }
      }
      const requestId = makeId('REQ');
      const now = nowIso();
      appendObjectRow_(reqSheet, reqMap, { RequestId: requestId, SenderEmail: senderEmail, ReceiverEmail: receiverEmail, Status: 'pending', Attempts: 1, CreatedAt: now, UpdatedAt: now });
      return jsonResponse({ status: 'success', msg: 'تم إرسال الطلب.', requestStatus: 'pending', attempts: 1, remainingAttempts: MAX_INTEREST_ATTEMPTS - 1, requestId: requestId });
    }

    // getMyRequests
    if (action === 'getMyRequests') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      const incoming = [], outgoing = [];
      for (let i = 1; i < reqRows.length; i++) {
        const senderEmail = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'SenderEmail'));
        const receiverEmail = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'ReceiverEmail'));
        const attempts = Number(getValueByHeader_(reqRows[i], reqMap, 'Attempts') || 1);
        const item = {
          requestId: getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '',
          senderEmail: senderEmail,
          receiverEmail: receiverEmail,
          status: String(getValueByHeader_(reqRows[i], reqMap, 'Status') || 'pending').toLowerCase(),
          attempts: attempts,
          remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - attempts),
          createdAt: getValueByHeader_(reqRows[i], reqMap, 'CreatedAt') || '',
          updatedAt: getValueByHeader_(reqRows[i], reqMap, 'UpdatedAt') || '',
          senderProfile: getPublicProfile_(profileRows, profileMap, senderEmail),
          receiverProfile: getPublicProfile_(profileRows, profileMap, receiverEmail)
        };
        if (receiverEmail === email) incoming.push(item);
        if (senderEmail === email) outgoing.push(item);
      }
      incoming.sort(function(a,b){ return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
      outgoing.sort(function(a,b){ return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
      return jsonResponse({ status: 'success', incoming: incoming, outgoing: outgoing });
    }

    // respondInterest
    if (action === 'respondInterest') {
      const requestId = String(data.requestId || '').trim();
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const responseStatus = String(data.responseStatus || '').trim().toLowerCase();
      if (responseStatus !== 'accepted' && responseStatus !== 'rejected') return jsonResponse({ status: 'error', msg: 'حالة غير صالحة.' });
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      for (let i = 1; i < reqRows.length; i++) {
        const id = String(getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '').trim();
        const receiverEmail = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'ReceiverEmail'));
        if (id === requestId) {
          if (receiverEmail !== email) return jsonResponse({ status: 'error', msg: 'غير مسموح لك بهذا الطلب.' });
          setValueByHeader_(reqSheet, i + 1, reqMap, 'Status', responseStatus);
          setValueByHeader_(reqSheet, i + 1, reqMap, 'UpdatedAt', nowIso());
          return jsonResponse({ status: 'success', requestStatus: responseStatus });
        }
      }
      return jsonResponse({ status: 'error', msg: 'الطلب غير موجود.' });
    }

    // getInterestStatus
    if (action === 'getInterestStatus') {
      const senderEmail = normalizeEmail(data.senderEmail);
      const receiverEmail = normalizeEmail(data.receiverEmail);
      if (!senderEmail || !receiverEmail) return jsonResponse({ status: 'success', requestStatus: 'none', attempts: 0, remainingAttempts: MAX_INTEREST_ATTEMPTS });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      let reverseAccepted = null;
      for (let i = 1; i < reqRows.length; i++) {
        const s = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'SenderEmail'));
        const r = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'ReceiverEmail'));
        const attempts = Number(getValueByHeader_(reqRows[i], reqMap, 'Attempts') || 1);
        const status = String(getValueByHeader_(reqRows[i], reqMap, 'Status') || 'pending').toLowerCase();
        const requestId = getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '';
        if (s === senderEmail && r === receiverEmail) {
          return jsonResponse({ status: 'success', requestStatus: status, requestId: requestId, attempts: attempts, remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - attempts) });
        }
        if (s === receiverEmail && r === senderEmail && status === 'accepted') {
          reverseAccepted = { status: 'success', requestStatus: 'accepted', requestId: requestId, attempts: attempts, remainingAttempts: Math.max(0, MAX_INTEREST_ATTEMPTS - attempts) };
        }
      }
      if (reverseAccepted) return jsonResponse(reverseAccepted);
      return jsonResponse({ status: 'success', requestStatus: 'none', attempts: 0, remainingAttempts: MAX_INTEREST_ATTEMPTS });
    }

    // requestAdminChat
    if (action === 'requestAdminChat') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      if (isAdmin(email)) return jsonResponse({ status: 'error', msg: 'الأدمن لا يحتاج إلى طلب.' });
      const sheet = getAdminRequestsSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (normalizeEmail(getValueByHeader_(rows[i], map, 'UserEmail')) === email) {
          return jsonResponse({ status: 'success', requestId: getValueByHeader_(rows[i], map, 'RequestId') || '', requestStatus: String(getValueByHeader_(rows[i], map, 'Status') || 'pending').toLowerCase(), remainingUserMessages: Number(getValueByHeader_(rows[i], map, 'RemainingUserMessages') || ADMIN_CHAT_USER_LIMIT) });
        }
      }
      const requestId = makeId('ADMREQ');
      const now = nowIso();
      appendObjectRow_(sheet, map, { RequestId: requestId, UserEmail: email, Status: 'pending', RemainingUserMessages: ADMIN_CHAT_USER_LIMIT, CreatedAt: now, UpdatedAt: now });
      return jsonResponse({ status: 'success', requestId: requestId, requestStatus: 'pending', remainingUserMessages: ADMIN_CHAT_USER_LIMIT });
    }

    // getAdminChatStatus
    if (action === 'getAdminChatStatus') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const sheet = getAdminRequestsSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (normalizeEmail(getValueByHeader_(rows[i], map, 'UserEmail')) === email) {
          return jsonResponse({ status: 'success', requestId: getValueByHeader_(rows[i], map, 'RequestId') || '', requestStatus: String(getValueByHeader_(rows[i], map, 'Status') || 'pending').toLowerCase(), remainingUserMessages: Number(getValueByHeader_(rows[i], map, 'RemainingUserMessages') || ADMIN_CHAT_USER_LIMIT) });
        }
      }
      return jsonResponse({ status: 'success', requestStatus: 'none', remainingUserMessages: ADMIN_CHAT_USER_LIMIT });
    }

    // adminGetChatRequests
    if (action === 'adminGetChatRequests') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const sheet = getAdminRequestsSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      const requests = [];
      for (let i = 1; i < rows.length; i++) {
        const userEmail = normalizeEmail(getValueByHeader_(rows[i], map, 'UserEmail'));
        requests.push({
          requestId: getValueByHeader_(rows[i], map, 'RequestId') || '',
          userEmail: userEmail,
          requestStatus: String(getValueByHeader_(rows[i], map, 'Status') || 'pending').toLowerCase(),
          remainingUserMessages: Number(getValueByHeader_(rows[i], map, 'RemainingUserMessages') || ADMIN_CHAT_USER_LIMIT),
          createdAt: getValueByHeader_(rows[i], map, 'CreatedAt') || '',
          updatedAt: getValueByHeader_(rows[i], map, 'UpdatedAt') || '',
          userProfile: getPublicProfile_(profileRows, profileMap, userEmail)
        });
      }
      requests.sort(function(a,b){ return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
      return jsonResponse({ status: 'success', requests: requests });
    }

    // adminOpenChatRequest
    if (action === 'adminOpenChatRequest') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const requestId = String(data.requestId || '').trim();
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const sheet = getAdminRequestsSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(getValueByHeader_(rows[i], map, 'RequestId') || '').trim() === requestId) {
          setValueByHeader_(sheet, i + 1, map, 'Status', 'opened');
          setValueByHeader_(sheet, i + 1, map, 'UpdatedAt', nowIso());
          return jsonResponse({ status: 'success', requestStatus: 'opened' });
        }
      }
      return jsonResponse({ status: 'error', msg: 'الطلب غير موجود.' });
    }

    // sendAdminChatMessage
    if (action === 'sendAdminChatMessage') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const requestId = String(data.requestId || '').trim();
      const text = cleanText(data.messageText);
      if (!text) return jsonResponse({ status: 'error', msg: 'الرسالة فارغة.' });
      if (text.length > MAX_MESSAGE_LENGTH) return jsonResponse({ status: 'error', msg: 'الرسالة طويلة جداً.' });
      const authUser = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!authUser.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const reqSheet = getAdminRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      for (let i = 1; i < reqRows.length; i++) {
        if (String(getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '').trim() === requestId) {
          const userEmail = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'UserEmail'));
          const status = String(getValueByHeader_(reqRows[i], reqMap, 'Status') || 'pending').toLowerCase();
          const remaining = Number(getValueByHeader_(reqRows[i], reqMap, 'RemainingUserMessages') || ADMIN_CHAT_USER_LIMIT);
          if (status === 'pending' && !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'الطلب مازال في طور المعالجة.' });
          if (!isAdmin(email) && userEmail !== email) return jsonResponse({ status: 'error', msg: 'غير مصرح لك بهذه المحادثة.' });
          const senderType = isAdmin(email) ? 'admin' : 'user';
          if (senderType === 'user') {
            if (remaining <= 0) return jsonResponse({ status: 'error', msg: 'لقد استهلكت 5 رسائل، انتظر رد الموجه.' });
            setValueByHeader_(reqSheet, i + 1, reqMap, 'RemainingUserMessages', remaining - 1);
          } else {
            setValueByHeader_(reqSheet, i + 1, reqMap, 'Status', 'opened');
            setValueByHeader_(reqSheet, i + 1, reqMap, 'RemainingUserMessages', ADMIN_CHAT_USER_LIMIT);
          }
          setValueByHeader_(reqSheet, i + 1, reqMap, 'UpdatedAt', nowIso());
          const msgSheet = getAdminMessagesSheet_();
          const msgMap = getHeaderMap_(msgSheet);
          appendObjectRow_(msgSheet, msgMap, { MessageId: makeId('ADMMSG'), RequestId: requestId, UserEmail: userEmail, SenderType: senderType, SenderEmail: email, MessageText: text, CreatedAt: nowIso() });
          return jsonResponse({ status: 'success', remainingUserMessages: senderType === 'admin' ? ADMIN_CHAT_USER_LIMIT : Math.max(0, remaining - 1) });
        }
      }
      return jsonResponse({ status: 'error', msg: 'طلب المحادثة غير موجود.' });
    }

    // getAdminChatMessages
    if (action === 'getAdminChatMessages') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const requestId = String(data.requestId || '').trim();
      const authUser = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!authUser.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const reqSheet = getAdminRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      let requestInfo = null;
      for (let i = 1; i < reqRows.length; i++) {
        if (String(getValueByHeader_(reqRows[i], reqMap, 'RequestId') || '').trim() === requestId) {
          requestInfo = {
            userEmail: normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'UserEmail')),
            requestStatus: String(getValueByHeader_(reqRows[i], reqMap, 'Status') || 'pending').toLowerCase(),
            remainingUserMessages: Number(getValueByHeader_(reqRows[i], reqMap, 'RemainingUserMessages') || ADMIN_CHAT_USER_LIMIT)
          };
          break;
        }
      }
      if (!requestInfo) return jsonResponse({ status: 'error', msg: 'الطلب غير موجود.' });
      if (!isAdmin(email) && requestInfo.userEmail !== email) return jsonResponse({ status: 'error', msg: 'غير مصرح لك بهذه المحادثة.' });
      const sheet = getAdminMessagesSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      const messages = [];
      for (let i = 1; i < rows.length; i++) {
        if (String(getValueByHeader_(rows[i], map, 'RequestId') || '').trim() === requestId) {
          messages.push({
            messageId: getValueByHeader_(rows[i], map, 'MessageId') || '',
            requestId: getValueByHeader_(rows[i], map, 'RequestId') || '',
            userEmail: getValueByHeader_(rows[i], map, 'UserEmail') || '',
            senderType: getValueByHeader_(rows[i], map, 'SenderType') || '',
            senderEmail: getValueByHeader_(rows[i], map, 'SenderEmail') || '',
            messageText: getValueByHeader_(rows[i], map, 'MessageText') || '',
            createdAt: getValueByHeader_(rows[i], map, 'CreatedAt') || ''
          });
        }
      }
      messages.sort(function(a,b){ return String(a.createdAt).localeCompare(String(b.createdAt)); });
      return jsonResponse({ status: 'success', requestInfo: requestInfo, messages: messages });
    }

    // sendMessage — v2.3: admin intervention allowed
    if (action === 'sendMessage') {
      const senderEmail = normalizeEmail(data.senderEmail);
      const receiverEmail = normalizeEmail(data.receiverEmail);
      const code = normalizeCode(data.code);
      const text = cleanText(data.messageText);
      if (!text) return jsonResponse({ status: 'error', msg: 'الرسالة فارغة.' });
      if (text.length > MAX_MESSAGE_LENGTH) return jsonResponse({ status: 'error', msg: 'الرسالة طويلة جداً.' });
      if (!senderEmail || !receiverEmail || senderEmail === receiverEmail) return jsonResponse({ status: 'error', msg: 'بيانات المحادثة غير صالحة.' });
      const senderIsAdmin = isAdmin(senderEmail);
      if (!senderIsAdmin && isAdmin(receiverEmail)) return jsonResponse({ status: 'error', msg: 'استعمل قسم الموجه بدل المحادثات الخاصة.' });
      const sender = findUserByEmailAndCode_(profileRows, profileMap, senderEmail, code);
      if (!sender.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const receiver = findUserByEmail_(profileRows, profileMap, receiverEmail);
      if (!receiver.found) return jsonResponse({ status: 'error', msg: 'المستخدم غير موجود.' });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      if (!senderIsAdmin && !canUsersChat_(reqRows, reqMap, senderEmail, receiverEmail)) return jsonResponse({ status: 'error', msg: 'المحادثة غير متاحة إلا بعد قبول الطلب.' });
      const pair = pairKey_(senderEmail, receiverEmail);
      const msgSheet = getMessagesSheet_();
      const msgMap = getHeaderMap_(msgSheet);
      appendObjectRow_(msgSheet, msgMap, { MessageId: makeId('MSG'), UserA: pair.userA, UserB: pair.userB, SenderEmail: senderEmail, MessageText: text, CreatedAt: nowIso(), ReadStatus: 'unread' });
      return jsonResponse({ status: 'success' });
    }

    // getMessages — v2.3: admin can monitor silently
    if (action === 'getMessages') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const otherEmail = normalizeEmail(data.otherEmail);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      const adminOverride = isAdmin(email);
      if (!adminOverride && !canUsersChat_(reqRows, reqMap, email, otherEmail)) return jsonResponse({ status: 'error', msg: 'لا توجد محادثة متاحة.' });
      const pair = pairKey_(email, otherEmail);
      const msgSheet = getMessagesSheet_();
      const msgMap = getHeaderMap_(msgSheet);
      const rows = msgSheet.getDataRange().getValues();
      const messages = [];
      for (let i = 1; i < rows.length; i++) {
        if (normalizeEmail(getValueByHeader_(rows[i], msgMap, 'UserA')) === pair.userA && normalizeEmail(getValueByHeader_(rows[i], msgMap, 'UserB')) === pair.userB) {
          messages.push({
            messageId: getValueByHeader_(rows[i], msgMap, 'MessageId') || '',
            senderEmail: getValueByHeader_(rows[i], msgMap, 'SenderEmail') || '',
            messageText: getValueByHeader_(rows[i], msgMap, 'MessageText') || '',
            createdAt: getValueByHeader_(rows[i], msgMap, 'CreatedAt') || '',
            readStatus: getValueByHeader_(rows[i], msgMap, 'ReadStatus') || ''
          });
        }
      }
      messages.sort(function(a,b){ return String(a.createdAt).localeCompare(String(b.createdAt)); });
      return jsonResponse({ status: 'success', messages: messages });
    }

    // getConversations
    if (action === 'getConversations') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found) return jsonResponse({ status: 'error', msg: 'تعذر التحقق من الحساب.' });
      if (isAdmin(email)) return jsonResponse({ status: 'success', conversations: [], isAdmin: true });
      const reqSheet = getRequestsSheet_();
      const reqMap = getHeaderMap_(reqSheet);
      const reqRows = reqSheet.getDataRange().getValues();
      const msgSheet = getMessagesSheet_();
      const msgMap = getHeaderMap_(msgSheet);
      const msgRows = msgSheet.getDataRange().getValues();
      const map = {};
      for (let i = 1; i < reqRows.length; i++) {
        const sender = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'SenderEmail'));
        const receiver = normalizeEmail(getValueByHeader_(reqRows[i], reqMap, 'ReceiverEmail'));
        const status = String(getValueByHeader_(reqRows[i], reqMap, 'Status') || '').toLowerCase();
        if (status !== 'accepted') continue;
        if (sender !== email && receiver !== email) continue;
        const pair = pairKey_(sender, receiver);
        const key = pair.userA + '||' + pair.userB;
        map[key] = { userA: pair.userA, userB: pair.userB, lastMessage: '', createdAt: getValueByHeader_(reqRows[i], reqMap, 'UpdatedAt') || getValueByHeader_(reqRows[i], reqMap, 'CreatedAt') || '' };
      }
      for (let i = 1; i < msgRows.length; i++) {
        const a = normalizeEmail(getValueByHeader_(msgRows[i], msgMap, 'UserA'));
        const b = normalizeEmail(getValueByHeader_(msgRows[i], msgMap, 'UserB'));
        if (a !== email && b !== email) continue;
        const key = a + '||' + b;
        if (!map[key]) map[key] = { userA: a, userB: b, lastMessage: '', createdAt: getValueByHeader_(msgRows[i], msgMap, 'CreatedAt') || '' };
        map[key].lastMessage = getValueByHeader_(msgRows[i], msgMap, 'MessageText') || '';
        map[key].createdAt = getValueByHeader_(msgRows[i], msgMap, 'CreatedAt') || map[key].createdAt;
      }
      const conversations = Object.keys(map).map(function(key){
        const item = map[key];
        const otherEmail = item.userA === email ? item.userB : item.userA;
        return { userA: item.userA, userB: item.userB, otherEmail: otherEmail, otherProfile: getPublicProfile_(profileRows, profileMap, otherEmail), lastMessage: item.lastMessage, createdAt: item.createdAt };
      }).sort(function(a,b){ return String(b.createdAt).localeCompare(String(a.createdAt)); });
      return jsonResponse({ status: 'success', conversations: conversations, isAdmin: false });
    }

    // adminReviewRegistrationRequest
    if (action === 'adminReviewRegistrationRequest') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const requestId = String(data.requestId || '').trim();
      const decision = String(data.decision || '').trim().toLowerCase();
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      if (decision !== 'approved' && decision !== 'rejected') return jsonResponse({ status: 'error', msg: 'قرار غير صالح.' });
      return jsonResponse(reviewRegistrationRequest_(requestId, decision, email));
    }

    // adminDashboard
    if (action === 'adminDashboard') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const requests = getRequestsSheet_().getDataRange().getValues();
      const ratings = getRatingsSheet_().getDataRange().getValues();
      const messages = getMessagesSheet_().getDataRange().getValues();
      const adminRequests = getAdminRequestsSheet_().getDataRange().getValues();
      const adminMessages = getAdminMessagesSheet_().getDataRange().getValues();
      const registrationSheet = getRegistrationRequestsSheet_();
      const registrationMap = getHeaderMap_(registrationSheet);
      const registrationRows = registrationSheet.getDataRange().getValues();
      const guideProfile = findUserByEmail_(profileRows, profileMap, email);
      const pendingRegistrationRequests = [];
      for (let i = 1; i < registrationRows.length; i++) {
        const requestStatus = String(getValueByHeader_(registrationRows[i], registrationMap, 'RequestStatus') || 'pending').toLowerCase();
        if (requestStatus !== 'pending') continue;
        const userEmail = normalizeEmail(getValueByHeader_(registrationRows[i], registrationMap, 'Email'));
        pendingRegistrationRequests.push({
          requestId: getValueByHeader_(registrationRows[i], registrationMap, 'RequestId') || '',
          requestStatus: requestStatus,
          email: userEmail,
          name: getValueByHeader_(registrationRows[i], registrationMap, 'Name') || '',
          age: getValueByHeader_(registrationRows[i], registrationMap, 'Age') || '',
          gender: getValueByHeader_(registrationRows[i], registrationMap, 'Gender') || '',
          statusText: getValueByHeader_(registrationRows[i], registrationMap, 'Status') || '',
          location: getValueByHeader_(registrationRows[i], registrationMap, 'Location') || '',
          birthCity: getValueByHeader_(registrationRows[i], registrationMap, 'BirthCity') || '',
          height: getValueByHeader_(registrationRows[i], registrationMap, 'Height') || '',
          imageUrl: getValueByHeader_(registrationRows[i], registrationMap, 'ImageURL') || '',
          specs: getValueByHeader_(registrationRows[i], registrationMap, 'Specs') || '',
          createdAt: getValueByHeader_(registrationRows[i], registrationMap, 'CreatedAt') || '',
          updatedAt: getValueByHeader_(registrationRows[i], registrationMap, 'UpdatedAt') || ''
        });
      }
      pendingRegistrationRequests.sort(function(a,b){ return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
      return jsonResponse({
        status: 'success',
        guideProfile: {
          name: GUIDE_NAME,
          label: GUIDE_LABEL,
          imageUrl: (guideProfile.found ? (getValueByHeader_(guideProfile.row, profileMap, 'ImageURL') || '') : '') || GUIDE_DEFAULT_IMAGE
        },
        totals: {
          profiles: Math.max(0, profileRows.length - 1),
          requests: Math.max(0, requests.length - 1),
          ratings: Math.max(0, ratings.length - 1),
          memberMessages: Math.max(0, messages.length - 1),
          adminChatRequests: Math.max(0, adminRequests.length - 1),
          adminMessages: Math.max(0, adminMessages.length - 1),
          pendingRegistrationRequests: pendingRegistrationRequests.length
        },
        registrationRequests: pendingRegistrationRequests
      });
    }

    // adminDeleteUser
    if (action === 'adminDeleteUser') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const targetEmail = normalizeEmail(data.targetEmail);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const target = findUserByEmail_(profileRows, profileMap, targetEmail);
      if (!target.found) return jsonResponse({ status: 'error', msg: 'الحساب غير موجود.' });
      if (isAdmin(targetEmail)) return jsonResponse({ status: 'error', msg: 'لا يمكن حذف حساب الموجه من هنا.' });
      deleteUserAndRelatedData_(targetEmail, target.rowIndex);
      markDeletedUser_(targetEmail, email, 'admin_delete');
      return jsonResponse({ status: 'success' });
    }

    // adminDeleteMessage
    if (action === 'adminDeleteMessage') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const messageId = String(data.messageId || '').trim();
      const scope = String(data.scope || 'member').trim().toLowerCase();
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const sheet = scope === 'admin' ? getAdminMessagesSheet_() : getMessagesSheet_();
      const map = getHeaderMap_(sheet);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(getValueByHeader_(rows[i], map, 'MessageId') || '').trim() === messageId) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ status: 'success' });
        }
      }
      return jsonResponse({ status: 'error', msg: 'الرسالة غير موجودة.' });
    }


    // ===== v2.3 — مراقبة الموجه الخفية + تدخل حي =====
    if (action === 'adminListMemberConversations') {
      const email = normalizeEmail(data.email);
      const code = normalizeCode(data.code);
      const user = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!user.found || !isAdmin(email)) return jsonResponse({ status: 'error', msg: 'غير مصرح لك.' });
      const msgSheet = getMessagesSheet_();
      const msgMap = getHeaderMap_(msgSheet);
      const msgRows = msgSheet.getDataRange().getValues();
      const convMap = {};
      for (let i = 1; i < msgRows.length; i++) {
        const a = normalizeEmail(getValueByHeader_(msgRows[i], msgMap, 'UserA'));
        const b = normalizeEmail(getValueByHeader_(msgRows[i], msgMap, 'UserB'));
        const text = getValueByHeader_(msgRows[i], msgMap, 'MessageText') || '';
        const createdAt = getValueByHeader_(msgRows[i], msgMap, 'CreatedAt') || '';
        const key = a + '||' + b;
        if (!convMap[key]) convMap[key] = { userA: a, userB: b, lastMessage: '', createdAt: '', count:0 };
        convMap[key].lastMessage = text; convMap[key].createdAt = createdAt; convMap[key].count++;
      }
      const list = Object.keys(convMap).map(function(k){
        const c = convMap[k];
        return { userA: c.userA, userB: c.userB, profileA: getPublicProfile_(profileRows, profileMap, c.userA), profileB: getPublicProfile_(profileRows, profileMap, c.userB), lastMessage: c.lastMessage, createdAt: c.createdAt, messageCount: c.count };
      }).sort(function(a,b){ return String(b.createdAt).localeCompare(String(a.createdAt)); });
      return jsonResponse({ status:'success', conversations: list });
    }
    if (action === 'adminGetMemberMessages') {
      const email = normalizeEmail(data.email); const code = normalizeCode(data.code);
      const userA = normalizeEmail(data.userA); const userB = normalizeEmail(data.userB);
      const auth = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!auth.found || !isAdmin(email)) return jsonResponse({ status:'error', msg:'غير مصرح لك.'});
      const pair = pairKey_(userA, userB);
      const msgSheet = getMessagesSheet_(); const msgMap = getHeaderMap_(msgSheet); const rows = msgSheet.getDataRange().getValues();
      const messages = [];
      for (let i=1;i<rows.length;i++){
        if (normalizeEmail(getValueByHeader_(rows[i], msgMap, 'UserA'))===pair.userA && normalizeEmail(getValueByHeader_(rows[i], msgMap, 'UserB'))===pair.userB){
          const senderEmail = getValueByHeader_(rows[i], msgMap, 'SenderEmail')||'';
          messages.push({ messageId: getValueByHeader_(rows[i], msgMap, 'MessageId')||'', senderEmail: senderEmail, senderProfile: getPublicProfile_(profileRows, profileMap, senderEmail), isAdminMsg: isAdmin(senderEmail), messageText: getValueByHeader_(rows[i], msgMap, 'MessageText')||'', createdAt: getValueByHeader_(rows[i], msgMap, 'CreatedAt')||'', readStatus: getValueByHeader_(rows[i], msgMap, 'ReadStatus')||'' });
        }
      }
      messages.sort(function(a,b){ return String(a.createdAt).localeCompare(String(b.createdAt)); });
      return jsonResponse({ status:'success', messages: messages, pair: pair, profileA: getPublicProfile_(profileRows, profileMap, pair.userA), profileB: getPublicProfile_(profileRows, profileMap, pair.userB) });
    }
    if (action === 'requestGuideIntervention') {
      const email = normalizeEmail(data.email); const code = normalizeCode(data.code);
      const otherEmail = normalizeEmail(data.otherEmail); const reason = cleanText(data.reason || 'طلب تدخل الموجه في المحادثة');
      const auth = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!auth.found) return jsonResponse({ status:'error', msg:'تعذر التحقق من الحساب.'});
      const sheet = getAdminRequestsSheet_(); const map = getHeaderMap_(sheet); const requestId = makeId('INTERV'); const now = nowIso();
      appendObjectRow_(sheet, map, { RequestId: requestId, UserEmail: email, Status: 'intervention:'+ pairKey_(email, otherEmail).userA + '<>' + pairKey_(email, otherEmail).userB, RemainingUserMessages: 0, CreatedAt: now, UpdatedAt: now });
      const adminMsgSheet = getAdminMessagesSheet_(); const adminMsgMap = getHeaderMap_(adminMsgSheet);
      appendObjectRow_(adminMsgSheet, adminMsgMap, { MessageId: makeId('SYS'), RequestId: requestId, UserEmail: email, SenderType: 'system', SenderEmail: 'system@matchnest', MessageText: 'طلب تدخل الموجه في محادثة بين ' + email + ' و ' + otherEmail + ' – السبب: ' + reason, CreatedAt: now });
      return jsonResponse({ status:'success', requestId: requestId, msg:'تم إرسال طلب تدخل الموجه. سيظهر الموجه في المحادثة قريباً.' });
    }
    if (action === 'adminSendInterventionMessage') {
      const email = normalizeEmail(data.email); const code = normalizeCode(data.code);
      const userA = normalizeEmail(data.userA); const userB = normalizeEmail(data.userB);
      const text = cleanText(data.messageText);
      const auth = findUserByEmailAndCode_(profileRows, profileMap, email, code);
      if (!auth.found || !isAdmin(email)) return jsonResponse({ status:'error', msg:'غير مصرح لك.'});
      if (!text) return jsonResponse({ status:'error', msg:'الرسالة فارغة.'});
      const pair = pairKey_(userA, userB);
      const msgSheet = getMessagesSheet_(); const msgMap = getHeaderMap_(msgSheet);
      appendObjectRow_(msgSheet, msgMap, { MessageId: makeId('GUIDE'), UserA: pair.userA, UserB: pair.userB, SenderEmail: email, MessageText: '🧭 الموجه: ' + text, CreatedAt: nowIso(), ReadStatus: 'unread' });
      return jsonResponse({ status:'success' });
    }


        return jsonResponse({ status: 'error', msg: 'طلب غير معروف: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', msg: 'خطأ داخلي في الخادم: ' + String(err) });
  }
}

function doPost(e) {
  // supports both GET and POST (JSON body)
  if (e && e.postData && e.postData.type === 'application/json') {
    try {
      e.parameter = JSON.parse(e.postData.contents);
    } catch(err){}
  }
  return doGet(e);
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}
