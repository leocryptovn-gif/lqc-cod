// ============================================================
// Netlify Serverless Function — COD Management API
// Auth: JWT + bcrypt | Roles: admin, manager
// ============================================================
const { google } = require('googleapis');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');

const DATA_SHEET  = 'DuLieu';
const RIDER_SHEET = 'Riders';
const USERS_SHEET = 'Users';

const JWT_SECRET  = () => process.env.JWT_SECRET || 'cod-secret-change-me';
const JWT_EXPIRES = '10h';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Content-Type': 'application/json',
};

const ok   = d    => ({ statusCode: 200, headers: CORS, body: JSON.stringify(d) });
const fail = (m, c=400) => ({ statusCode: c, headers: CORS, body: JSON.stringify({ error: m }) });

function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}
const getSheets = () => google.sheets({ version: 'v4', auth: getAuth() });
const SID = () => process.env.SPREADSHEET_ID;

function verifyToken(token) {
  try { return jwt.verify(token, JWT_SECRET()); } catch { return null; }
}
function getToken(event) {
  const h = event.headers.authorization || event.headers.Authorization || '';
  return h.startsWith('Bearer ') ? h.slice(7) : null;
}

// ============================================================
// INIT SHEETS
// ============================================================
async function ensureSheets(sheets) {
  const ss = await sheets.spreadsheets.get({ spreadsheetId: SID() });
  const existing = ss.data.sheets.map(s => s.properties.title);
  const toAdd = [DATA_SHEET, RIDER_SHEET, USERS_SHEET].filter(n => !existing.includes(n));

  if (toAdd.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SID(),
      requestBody: { requests: toAdd.map(t => ({ addSheet: { properties: { title: t } } })) },
    });
  }

  if (!existing.includes(DATA_SHEET)) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SID(), range: `${DATA_SHEET}!A1:F1`, valueInputOption: 'RAW',
      requestBody: { values: [['Ngay', 'Rider', 'So kien', 'So tien COD', 'Ghi chu', 'Thoi gian nhap']] },
    });
  }

  if (!existing.includes(RIDER_SHEET)) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SID(), range: `${RIDER_SHEET}!A1:A4`, valueInputOption: 'RAW',
      requestBody: { values: [['Ten Rider'], ['Rider A'], ['Rider B'], ['Rider C']] },
    });
  }

  if (!existing.includes(USERS_SHEET)) {
    const hash = bcrypt.hashSync('Admin@123', 10);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SID(), range: `${USERS_SHEET}!A1:E2`, valueInputOption: 'RAW',
      requestBody: { values: [
        ['Username', 'PasswordHash', 'Role', 'HoTen', 'Active'],
        ['admin', hash, 'admin', 'Administrator', 'TRUE'],
      ]},
    });
  }
}

// ============================================================
// AUTH
// ============================================================
async function loginUser(sheets, username, password) {
  if (!username || !password) return { error: 'Vui long nhap du thong tin.' };
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${USERS_SHEET}!A2:E` });
  const rows = res.data.values || [];
  const row = rows.find(r => r[0] === username);
  if (!row)           return { error: 'Tai khoan khong ton tai.' };
  if (row[4] !== 'TRUE') return { error: 'Tai khoan da bi khoa.' };
  if (!bcrypt.compareSync(password, row[1])) return { error: 'Mat khau khong dung.' };
  const payload = { username: row[0], role: row[2], fullName: row[3] };
  return { token: jwt.sign(payload, JWT_SECRET(), { expiresIn: JWT_EXPIRES }), user: payload };
}

// ============================================================
// RIDERS
// ============================================================
async function getRiders(sheets) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${RIDER_SHEET}!A2:A` });
  return (res.data.values || []).map(r => r[0]).filter(Boolean);
}

async function addRider(sheets, name) {
  const riders = await getRiders(sheets);
  if (riders.map(r => r.toLowerCase()).includes(name.trim().toLowerCase()))
    return { success: false, message: `Rider '${name}' da ton tai.` };
  await sheets.spreadsheets.values.append({
    spreadsheetId: SID(), range: `${RIDER_SHEET}!A:A`, valueInputOption: 'RAW',
    requestBody: { values: [[name.trim()]] },
  });
  return { success: true, message: `Da them rider '${name.trim()}'.` };
}

async function deleteRider(sheets, name) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${RIDER_SHEET}!A:A` });
  const rows = res.data.values || [];
  const idx = rows.findIndex((r, i) => i > 0 && r[0] === name);
  if (idx === -1) return { success: false, message: `Khong tim thay rider '${name}'.` };
  const ss2 = await sheets.spreadsheets.get({ spreadsheetId: SID() });
  const sid = ss2.data.sheets.find(s => s.properties.title === RIDER_SHEET).properties.sheetId;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID(),
    requestBody: { requests: [{ deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 } } }] },
  });
  return { success: true, message: `Da xoa rider '${name}'.` };
}

// ============================================================
// COD RECORDS
// ============================================================
async function getData(sheets, filter = {}) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${DATA_SHEET}!A2:F` });
  const result = [];
  (res.data.values || []).forEach((row, idx) => {
    if (!row[0]) return;
    const raw = row[0];
    let d;
    if (raw.includes('/')) { const p = raw.split('/'); d = new Date(+p[2], +p[1]-1, +p[0]); }
    else d = new Date(raw);
    if (!d || isNaN(d)) return;
    const dateSort = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    const dateDisplay = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    if (filter.startDate && dateSort < filter.startDate) return;
    if (filter.endDate   && dateSort > filter.endDate)   return;
    if (filter.rider && filter.rider !== '' && row[1] !== filter.rider) return;
    result.push({ rowIndex: idx+2, date: dateDisplay, dateSort, rider: row[1]||'', packages: parseInt(row[2])||0, amount: parseFloat(row[3])||0, note: row[4]||'', timestamp: row[5]||'' });
  });
  result.sort((a,b) => b.dateSort.localeCompare(a.dateSort) || b.rowIndex - a.rowIndex);
  return result;
}

async function addRecord(sheets, data) {
  if (!data.date || !data.rider || !data.packages || data.amount === undefined)
    return { success: false, message: 'Vui long dien day du thong tin.' };
  const ts = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SID(), range: `${DATA_SHEET}!A:F`, valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[data.date, data.rider, parseInt(data.packages), parseFloat(data.amount), data.note||'', ts]] },
  });
  return { success: true, message: 'Da luu thanh cong!' };
}

async function updateRecord(sheets, rowIndex, data) {
  const ts = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SID(), range: `${DATA_SHEET}!A${rowIndex}:F${rowIndex}`, valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[data.date, data.rider, parseInt(data.packages), parseFloat(data.amount), data.note||'', ts]] },
  });
  return { success: true, message: 'Da cap nhat ban ghi.' };
}

async function deleteRecord(sheets, rowIndex) {
  const ss2 = await sheets.spreadsheets.get({ spreadsheetId: SID() });
  const sid = ss2.data.sheets.find(s => s.properties.title === DATA_SHEET).properties.sheetId;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID(),
    requestBody: { requests: [{ deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: rowIndex-1, endIndex: rowIndex } } }] },
  });
  return { success: true, message: 'Da xoa ban ghi.' };
}

async function getSummary(sheets, filter = {}) {
  const data = await getData(sheets, filter);
  let totalCOD = 0, totalPackages = 0;
  const riderMap = {}, dateMap = {};
  data.forEach(r => {
    totalCOD += r.amount; totalPackages += r.packages;
    if (!riderMap[r.rider]) riderMap[r.rider] = { rider: r.rider, packages:0, amount:0, records:0 };
    riderMap[r.rider].packages += r.packages; riderMap[r.rider].amount += r.amount; riderMap[r.rider].records++;
    if (!dateMap[r.dateSort]) dateMap[r.dateSort] = { date: r.date, dateSort: r.dateSort, packages:0, amount:0, records:0 };
    dateMap[r.dateSort].packages += r.packages; dateMap[r.dateSort].amount += r.amount; dateMap[r.dateSort].records++;
  });
  return {
    totalCOD, totalPackages, totalRecords: data.length,
    byRider: Object.values(riderMap).sort((a,b) => b.amount - a.amount),
    byDate:  Object.values(dateMap).sort((a,b) => b.dateSort.localeCompare(a.dateSort)).slice(0,14),
  };
}

// ============================================================
// USER MANAGEMENT (admin only)
// ============================================================
async function listUsers(sheets) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${USERS_SHEET}!A2:E` });
  return (res.data.values || []).map((r, i) => ({
    rowIndex: i+2, username: r[0]||'', role: r[2]||'', fullName: r[3]||'', active: r[4]==='TRUE',
  }));
}

async function addUser(sheets, data) {
  const users = await listUsers(sheets);
  if (users.find(u => u.username === data.username))
    return { success: false, message: `Username '${data.username}' da ton tai.` };
  const hash = bcrypt.hashSync(data.password, 10);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SID(), range: `${USERS_SHEET}!A:E`, valueInputOption: 'RAW',
    requestBody: { values: [[data.username, hash, data.role, data.fullName, 'TRUE']] },
  });
  return { success: true, message: `Da tao tai khoan '${data.username}'.` };
}

async function updateUser(sheets, rowIndex, data) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SID(), range: `${USERS_SHEET}!C${rowIndex}:E${rowIndex}`, valueInputOption: 'RAW',
    requestBody: { values: [[data.role, data.fullName, data.active ? 'TRUE' : 'FALSE']] },
  });
  return { success: true, message: 'Da cap nhat tai khoan.' };
}

async function deleteUser(sheets, rowIndex) {
  const ss2 = await sheets.spreadsheets.get({ spreadsheetId: SID() });
  const sid = ss2.data.sheets.find(s => s.properties.title === USERS_SHEET).properties.sheetId;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID(),
    requestBody: { requests: [{ deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: rowIndex-1, endIndex: rowIndex } } }] },
  });
  return { success: true, message: 'Da xoa tai khoan.' };
}

async function resetPassword(sheets, rowIndex, newPassword) {
  const hash = bcrypt.hashSync(newPassword, 10);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SID(), range: `${USERS_SHEET}!B${rowIndex}`, valueInputOption: 'RAW',
    requestBody: { values: [[hash]] },
  });
  return { success: true, message: 'Da dat lai mat khau.' };
}

async function changeOwnPassword(sheets, username, currentPassword, newPassword) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SID(), range: `${USERS_SHEET}!A2:E` });
  const rows = res.data.values || [];
  const idx = rows.findIndex(r => r[0] === username);
  if (idx === -1) return { success: false, message: 'Tai khoan khong ton tai.' };
  if (!bcrypt.compareSync(currentPassword, rows[idx][1])) return { success: false, message: 'Mat khau hien tai khong dung.' };
  const hash = bcrypt.hashSync(newPassword, 10);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SID(), range: `${USERS_SHEET}!B${idx+2}`, valueInputOption: 'RAW',
    requestBody: { values: [[hash]] },
  });
  return { success: true, message: 'Da thay doi mat khau thanh cong.' };
}

// ============================================================
// MAIN HANDLER
// ============================================================
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: CORS, body: '' };
  try {
    const sheets = getSheets();
    await ensureSheets(sheets);

    // Login — no auth required
    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      if (body.action === 'login') return ok(await loginUser(sheets, body.username, body.password));
    }

    // All other endpoints require valid token
    const token = getToken(event);
    const user  = verifyToken(token);
    if (!user) return fail('Chua dang nhap hoac phien lam viec het han.', 401);
    const isAdmin = user.role === 'admin';

    if (event.httpMethod === 'GET') {
      const p = event.queryStringParameters || {};
      switch (p.action) {
        case 'getRiders':  return ok(await getRiders(sheets));
        case 'getData':    return ok(await getData(sheets, { startDate: p.startDate, endDate: p.endDate, rider: p.rider }));
        case 'getSummary': return ok(await getSummary(sheets, { startDate: p.startDate, endDate: p.endDate, rider: p.rider }));
        case 'listUsers':  if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await listUsers(sheets));
        default: return fail('Unknown GET action');
      }
    }

    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      switch (body.action) {
        case 'addRecord':       return ok(await addRecord(sheets, body.data));
        case 'updateRecord':    if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await updateRecord(sheets, body.rowIndex, body.data));
        case 'deleteRecord':    if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await deleteRecord(sheets, body.rowIndex));
        case 'addRider':        return ok(await addRider(sheets, body.name));
        case 'deleteRider':     if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await deleteRider(sheets, body.name));
        case 'addUser':         if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await addUser(sheets, body.data));
        case 'updateUser':      if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await updateUser(sheets, body.rowIndex, body.data));
        case 'deleteUser':      if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await deleteUser(sheets, body.rowIndex));
        case 'resetPassword':   if (!isAdmin) return fail('Khong co quyen.', 403); return ok(await resetPassword(sheets, body.rowIndex, body.newPassword));
        case 'changePassword':  return ok(await changeOwnPassword(sheets, user.username, body.currentPassword, body.newPassword));
        default: return fail('Unknown POST action');
      }
    }

    return fail('Method not allowed', 405);
  } catch (e) {
    console.error(e);
    return fail('Server error: ' + e.message, 500);
  }
};
