// ============================================================
// Netlify Serverless Function — COD Management API
// Giao tiếp với Google Sheets qua Service Account
// ============================================================

const { google } = require('googleapis');

// ENV vars (set in Netlify dashboard):
//   GOOGLE_SERVICE_ACCOUNT_KEY  — JSON string của service account key
//   SPREADSHEET_ID              — ID của Google Sheet

const DATA_SHEET = 'DuLieu';
const RIDER_SHEET = 'Riders';

// ------ Auth ------
function getAuth() {
  const keyJson = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  return new google.auth.GoogleAuth({
    credentials: keyJson,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

function getSheets() {
  return google.sheets({ version: 'v4', auth: getAuth() });
}

const SPREADSHEET_ID = () => process.env.SPREADSHEET_ID;

// ------ CORS headers ------
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Content-Type': 'application/json',
};

function ok(data) {
  return { statusCode: 200, headers: CORS, body: JSON.stringify(data) };
}
function err(msg, code = 400) {
  return { statusCode: code, headers: CORS, body: JSON.stringify({ error: msg }) };
}

// ------ Init sheets nếu chưa có ------
async function ensureSheets(sheets) {
  const ss = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID() });
  const sheetNames = ss.data.sheets.map(s => s.properties.title);

  if (!sheetNames.includes(DATA_SHEET)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID(),
      requestBody: {
        requests: [{ addSheet: { properties: { title: DATA_SHEET } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID(),
      range: `${DATA_SHEET}!A1:F1`,
      valueInputOption: 'RAW',
      requestBody: { values: [['Ngày', 'Rider', 'Số kiện', 'Số tiền COD', 'Ghi chú', 'Thời gian nhập']] },
    });
  }

  if (!sheetNames.includes(RIDER_SHEET)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID(),
      requestBody: {
        requests: [{ addSheet: { properties: { title: RIDER_SHEET } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID(),
      range: `${RIDER_SHEET}!A1:A4`,
      valueInputOption: 'RAW',
      requestBody: { values: [['Tên Rider'], ['Nguyễn Văn A'], ['Trần Thị B'], ['Lê Văn C']] },
    });
  }
}

// ------ getRiders ------
async function getRiders(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${RIDER_SHEET}!A2:A`,
  });
  const rows = res.data.values || [];
  return rows.map(r => r[0]).filter(Boolean);
}

// ------ addRider ------
async function addRider(sheets, name) {
  const riders = await getRiders(sheets);
  if (riders.map(r => r.toLowerCase()).includes(name.trim().toLowerCase())) {
    return { success: false, message: `Rider '${name}' đã tồn tại.` };
  }
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${RIDER_SHEET}!A:A`,
    valueInputOption: 'RAW',
    requestBody: { values: [[name.trim()]] },
  });
  return { success: true, message: `Đã thêm rider '${name.trim()}'.` };
}

// ------ deleteRider ------
async function deleteRider(sheets, name) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${RIDER_SHEET}!A:A`,
  });
  const rows = res.data.values || [];
  const idx = rows.findIndex((r, i) => i > 0 && r[0] === name);
  if (idx === -1) return { success: false, message: `Không tìm thấy rider '${name}'.` };

  // Lấy sheetId của RIDER_SHEET
  const ss = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID() });
  const sheet = ss.data.sheets.find(s => s.properties.title === RIDER_SHEET);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID(),
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 },
        },
      }],
    },
  });
  return { success: true, message: `Đã xóa rider '${name}'.` };
}

// ------ addRecord ------
async function addRecord(sheets, data) {
  if (!data.date || !data.rider || !data.packages || data.amount === undefined) {
    return { success: false, message: 'Vui lòng điền đầy đủ thông tin bắt buộc.' };
  }
  const timestamp = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${DATA_SHEET}!A:F`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[data.date, data.rider, parseInt(data.packages), parseFloat(data.amount), data.note || '', timestamp]],
    },
  });
  return { success: true, message: 'Đã lưu thành công!' };
}

// ------ getData ------
async function getData(sheets, filter = {}) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${DATA_SHEET}!A2:F`,
  });
  const rows = res.data.values || [];
  let result = [];

  rows.forEach((row, idx) => {
    if (!row[0]) return;
    // Parse ngày từ định dạng "d/M/yyyy" hoặc "yyyy-MM-dd"
    const rawDate = row[0];
    let dateObj;
    if (rawDate.includes('/')) {
      const parts = rawDate.split('/');
      // d/M/yyyy
      if (parts.length === 3) {
        dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    } else {
      dateObj = new Date(rawDate);
    }
    if (!dateObj || isNaN(dateObj)) return;

    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');
    const dateSort = `${yyyy}-${mm}-${dd}`;
    const dateDisplay = `${dd}/${mm}/${yyyy}`;

    if (filter.startDate && dateSort < filter.startDate) return;
    if (filter.endDate && dateSort > filter.endDate) return;
    if (filter.rider && filter.rider !== '' && row[1] !== filter.rider) return;

    result.push({
      rowIndex: idx + 2,
      date: dateDisplay,
      dateSort,
      rider: row[1] || '',
      packages: parseInt(row[2]) || 0,
      amount: parseFloat(row[3]) || 0,
      note: row[4] || '',
      timestamp: row[5] || '',
    });
  });

  result.sort((a, b) => b.dateSort.localeCompare(a.dateSort) || b.rowIndex - a.rowIndex);
  return result;
}

// ------ getSummary ------
async function getSummary(sheets, filter = {}) {
  const data = await getData(sheets, filter);
  let totalCOD = 0, totalPackages = 0;
  const riderMap = {}, dateMap = {};

  data.forEach(row => {
    totalCOD += row.amount;
    totalPackages += row.packages;
    if (!riderMap[row.rider]) riderMap[row.rider] = { rider: row.rider, packages: 0, amount: 0, records: 0 };
    riderMap[row.rider].packages += row.packages;
    riderMap[row.rider].amount += row.amount;
    riderMap[row.rider].records++;
    if (!dateMap[row.dateSort]) dateMap[row.dateSort] = { date: row.date, dateSort: row.dateSort, packages: 0, amount: 0, records: 0 };
    dateMap[row.dateSort].packages += row.packages;
    dateMap[row.dateSort].amount += row.amount;
    dateMap[row.dateSort].records++;
  });

  return {
    totalCOD,
    totalPackages,
    totalRecords: data.length,
    byRider: Object.values(riderMap).sort((a, b) => b.amount - a.amount),
    byDate: Object.values(dateMap).sort((a, b) => b.dateSort.localeCompare(a.dateSort)).slice(0, 14),
  };
}

// ------ deleteRecord ------
async function deleteRecord(sheets, rowIndex) {
  const ss = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID() });
  const sheet = ss.data.sheets.find(s => s.properties.title === DATA_SHEET);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID(),
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex },
        },
      }],
    },
  });
  return { success: true, message: 'Đã xóa bản ghi.' };
}

// ------ updateRecord ------
async function updateRecord(sheets, rowIndex, data) {
  const timestamp = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID(),
    range: `${DATA_SHEET}!A${rowIndex}:F${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[data.date, data.rider, parseInt(data.packages), parseFloat(data.amount), data.note || '', timestamp]],
    },
  });
  return { success: true, message: 'Đã cập nhật bản ghi.' };
}

// ============================================================
// MAIN HANDLER
// ============================================================
exports.handler = async (event) => {
  // CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS, body: '' };
  }

  try {
    const sheets = getSheets();
    await ensureSheets(sheets);

    if (event.httpMethod === 'GET') {
      const params = event.queryStringParameters || {};
      const action = params.action;

      if (action === 'getRiders') return ok(await getRiders(sheets));
      if (action === 'getData') return ok(await getData(sheets, {
        startDate: params.startDate || '',
        endDate: params.endDate || '',
        rider: params.rider || '',
      }));
      if (action === 'getSummary') return ok(await getSummary(sheets, {
        startDate: params.startDate || '',
        endDate: params.endDate || '',
        rider: params.rider || '',
      }));
      return err('Unknown action');
    }

    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      const action = body.action;

      if (action === 'addRecord') return ok(await addRecord(sheets, body.data));
      if (action === 'addRider') return ok(await addRider(sheets, body.name));
      if (action === 'deleteRider') return ok(await deleteRider(sheets, body.name));
      if (action === 'deleteRecord') return ok(await deleteRecord(sheets, body.rowIndex));
      if (action === 'updateRecord') return ok(await updateRecord(sheets, body.rowIndex, body.data));
      return err('Unknown action');
    }

    return err('Method not allowed', 405);
  } catch (e) {
    console.error(e);
    return err('Server error: ' + e.message, 500);
  }
};
