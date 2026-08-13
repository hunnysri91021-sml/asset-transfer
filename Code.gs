/**
 * SML Asset Transfer System — Google Apps Script Backend
 * ระบบโอนย้ายทรัพย์สิน — บริษัท สยามกลการโลจิสติกส์ จำกัด
 *
 * DEPLOY:
 *  1. เปิด Google Sheet ที่จะใช้เป็นฐานข้อมูล (สร้างใหม่ก็ได้)
 *  2. Extensions > Apps Script > วางไฟล์นี้ทับ Code.gs
 *  3. แก้ค่าคงที่ CONFIG ด้านล่าง (โดยเฉพาะ FRONTEND_URL หลัง deploy GitHub Pages แล้ว)
 *  4. Deploy > New deployment > Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5. คัดลอก Web App URL ไปใส่ใน index.html (GAS_URL)
 *  6. รันฟังก์ชัน setup() หนึ่งครั้งจากตัวแก้ไข Apps Script เพื่อสร้างชีตทั้งหมดอัตโนมัติ
 */

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  SHEET_ID: '', // เว้นว่าง = ใช้ Sheet ที่ผูกกับสคริปต์นี้ (แนะนำ) หรือใส่ Sheet ID ถ้ารันแบบ standalone
  FRONTEND_URL: 'https://hunnysri91021-sml.github.io/asset-transfer', // TODO: แก้เป็น URL จริงหลัง deploy GitHub Pages
  COMPANY_NAME: 'บริษัท สยามกลการโลจิสติกส์ จำกัด',
  DRIVE_FOLDER_NAME: 'SML_Asset_Transfer_Images',
  DEFAULT_TIMEOUT_MS: 15000
};

const SHEETS = {
  ASSETS: 'Assets',
  DEPT_CODES: 'DeptCodes',
  TRANSFERS: 'Transfers',
  ITEMS: 'TransferItems',
  LOG: 'ActivityLog'
};

const HEADERS = {
  ASSETS: ['AssetID', 'AssetName', 'Department', 'PurchaseDate', 'PurchasePrice', 'BookValue', 'Custodian', 'Location', 'ImageURL', 'ImageURLOverride', 'UpdatedAt'],
  DEPT_CODES: ['DeptName', 'Code'],
  TRANSFERS: ['TransferID', 'RunningNo', 'CreatedAt', 'Subject', 'SubjectOther', 'Purpose', 'FromDept', 'FromDeptCode', 'ToDept', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail'],
  ITEMS: ['TransferID', 'LineNo', 'AssetID', 'AssetName', 'FromDeptName', 'FromSignName', 'ToDeptName', 'ToSignName', 'Remark', 'ImageURL'],
  LOG: ['Timestamp', 'TransferID', 'Action', 'By', 'Detail']
};

const DEFAULT_DEPT_CODES = [
  ['Admin', 'ADM'],
  ['ตรวจสอบ', 'INS'],
  ['PAR', 'PAR'],
  ['ส่วนกลาง', 'CEN'],
  ['สำนักงาน PDI', 'PDI'],
  ['บัญชี', 'ACB'],
  ['อาคาร', 'BLD'],
  ['ล้างรถ', 'WAS'],
  ['ติดตั้ง ACC', 'ACC'],
  ['Safety', 'SAF'],
  ['Yard', 'YRD'],
  ['ปรับแต่ง', 'REP'],
  ['บุคคลและกิจการทั่วไป', 'ADM']
];

const STATUS = {
  DRAFT: 'Draft',
  PENDING: 'PendingApproval',
  APPROVED: 'Approved',
  REJECTED: 'Rejected'
};

// ============================================================
// SETUP — run once manually from the Apps Script editor
// ============================================================
function setup() {
  const ss = getSS_();
  ensureSheet_(ss, SHEETS.ASSETS, HEADERS.ASSETS);
  const deptSheet = ensureSheet_(ss, SHEETS.DEPT_CODES, HEADERS.DEPT_CODES);
  if (deptSheet.getLastRow() < 2) {
    deptSheet.getRange(2, 1, DEFAULT_DEPT_CODES.length, 2).setValues(DEFAULT_DEPT_CODES);
  }
  ensureSheet_(ss, SHEETS.TRANSFERS, HEADERS.TRANSFERS);
  ensureSheet_(ss, SHEETS.ITEMS, HEADERS.ITEMS);
  ensureSheet_(ss, SHEETS.LOG, HEADERS.LOG);
  Logger.log('Setup complete. Sheets ready: ' + Object.values(SHEETS).join(', '));
}

function getSS_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    // make sure header row matches (append any missing columns at the end)
    const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = headers.filter(h => existing.indexOf(h) === -1);
    if (missing.length) {
      sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sh;
}

// ============================================================
// HTTP ENTRY POINTS
// ============================================================
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    let result;
    switch (action) {
      case 'ping':
        result = { ok: true, time: new Date().toISOString() };
        break;
      case 'getAssets':
        result = { ok: true, data: getAssets_(e.parameter.q || '') };
        break;
      case 'getDeptCodes':
        result = { ok: true, data: getDeptCodes_() };
        break;
      case 'getTransfers':
        result = { ok: true, data: getTransfers_(e.parameter) };
        break;
      case 'getTransfer':
        result = { ok: true, data: getTransferFull_(e.parameter.id) };
        break;
      case 'getApprovalView':
        result = getApprovalView_(e.parameter.id, e.parameter.token);
        break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return respond_(result, e.parameter.callback);
  } catch (err) {
    return respond_({ ok: false, error: String(err) }, e.parameter.callback);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch (action) {
      case 'createTransfer':
        result = createTransfer_(body);
        break;
      case 'decideTransfer':
        result = decideTransfer_(body);
        break;
      case 'uploadImage':
        result = uploadImage_(body);
        break;
      case 'updateAssetImage':
        result = updateAssetImage_(body);
        break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return respond_(result);
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

function respond_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ASSETS
// ============================================================
function getAssets_(q) {
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  let rows = values.map(r => rowToObj_(r, idx));
  if (q) {
    const qq = q.toString().toLowerCase();
    rows = rows.filter(r =>
      String(r.AssetID).toLowerCase().indexOf(qq) !== -1 ||
      String(r.AssetName).toLowerCase().indexOf(qq) !== -1
    );
  }
  // resolve display image: override wins over original
  rows.forEach(r => { r.DisplayImage = r.ImageURLOverride || r.ImageURL || ''; });
  return rows;
}

function updateAssetImage_(body) {
  const assetId = String(body.assetId || '');
  if (!assetId) return { ok: false, error: 'assetId required' };
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.AssetID]) === assetId) {
      sh.getRange(i + 1, idx.ImageURLOverride + 1).setValue(body.imageUrl || '');
      sh.getRange(i + 1, idx.UpdatedAt + 1).setValue(new Date());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Asset not found: ' + assetId };
}

// ============================================================
// DEPT CODES
// ============================================================
function getDeptCodes_() {
  const sh = getSS_().getSheetByName(SHEETS.DEPT_CODES);
  const values = sh.getDataRange().getValues();
  values.shift();
  return values.filter(r => r[0]).map(r => ({ DeptName: r[0], Code: r[1] }));
}

// ============================================================
// TRANSFERS
// ============================================================
function getTransfers_(params) {
  const sh = getSS_().getSheetByName(SHEETS.TRANSFERS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  let rows = values.map(r => rowToObj_(r, idx));

  if (params.status) rows = rows.filter(r => r.Status === params.status);
  if (params.dept) rows = rows.filter(r => r.FromDept === params.dept || r.ToDept === params.dept);
  if (params.from) rows = rows.filter(r => new Date(r.CreatedAt) >= new Date(params.from));
  if (params.to) rows = rows.filter(r => new Date(r.CreatedAt) <= new Date(params.to + 'T23:59:59'));
  if (params.q) {
    const qq = params.q.toLowerCase();
    rows = rows.filter(r =>
      String(r.RunningNo).toLowerCase().indexOf(qq) !== -1 ||
      String(r.Purpose).toLowerCase().indexOf(qq) !== -1
    );
  }
  rows.sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));

  // attach item count
  const itemSh = getSS_().getSheetByName(SHEETS.ITEMS);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  const counts = {};
  itemValues.forEach(r => {
    const tid = r[itemIdx.TransferID];
    counts[tid] = (counts[tid] || 0) + 1;
  });
  rows.forEach(r => { r.ItemCount = counts[r.TransferID] || 0; });

  return rows;
}

function getTransferFull_(transferId) {
  const t = findTransferRow_(transferId);
  if (!t) return null;
  const items = getTransferItems_(transferId);
  const obj = t.obj;
  obj.Items = items;
  delete obj.ApprovalToken; // never leak token in general reads
  return obj;
}

function getTransferItems_(transferId) {
  const sh = getSS_().getSheetByName(SHEETS.ITEMS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  return values
    .map(r => rowToObj_(r, idx))
    .filter(r => r.TransferID === transferId)
    .sort((a, b) => a.LineNo - b.LineNo);
}

function findTransferRow_(transferId) {
  const sh = getSS_().getSheetByName(SHEETS.TRANSFERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.TransferID]) === String(transferId)) {
      return { rowNum: i + 1, obj: rowToObj_(values[i], idx), idx };
    }
  }
  return null;
}

function createTransfer_(body) {
  const items = body.items || [];
  if (!items.length) return { ok: false, error: 'ต้องมีรายการทรัพย์สินอย่างน้อย 1 รายการ' };
  if (!body.approverEmail) return { ok: false, error: 'ต้องระบุอีเมลผู้อนุมัติ' };

  const fromDeptCode = getCodeForDept_(body.fromDept) || 'GEN';
  const runningNo = getNextRunningNo_(fromDeptCode);
  const transferId = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();

  const tSheet = getSS_().getSheetByName(SHEETS.TRANSFERS);
  tSheet.appendRow([
    transferId,
    runningNo,
    now,
    body.subject || 'โอนย้าย',
    body.subjectOther || '',
    body.purpose || '',
    body.fromDept || '',
    fromDeptCode,
    body.toDept || '',
    STATUS.PENDING,
    body.approverName || '',
    body.approverEmail || '',
    token,
    '',
    '',
    body.createdBy || '',
    body.createdByEmail || ''
  ]);

  const iSheet = getSS_().getSheetByName(SHEETS.ITEMS);
  const itemRows = items.map((it, i) => [
    transferId,
    i + 1,
    it.assetId || '',
    it.assetName || '',
    it.fromDeptName || body.fromDept || '',
    it.fromSignName || '',
    it.toDeptName || body.toDept || '',
    it.toSignName || '',
    it.remark || '',
    it.imageUrl || ''
  ]);
  if (itemRows.length) {
    iSheet.getRange(iSheet.getLastRow() + 1, 1, itemRows.length, HEADERS.ITEMS.length).setValues(itemRows);
  }

  logActivity_(transferId, 'CREATE', body.createdBy || 'unknown', 'สร้างใบโอนย้าย ' + runningNo);

  const emailResult = sendApprovalEmail_(transferId, runningNo, body, items, token);

  return { ok: true, data: { transferId, runningNo, emailSent: emailResult.ok, emailError: emailResult.error || null } };
}

function getCodeForDept_(deptName) {
  const codes = getDeptCodes_();
  const found = codes.find(c => c.DeptName === deptName);
  return found ? found.Code : null;
}

function getNextRunningNo_(code) {
  const buddhistYear = new Date().getFullYear() + 543;
  const yy = String(buddhistYear).slice(-2);
  const sh = getSS_().getSheetByName(SHEETS.TRANSFERS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  const prefix = code + '/';
  const suffix = '/' + yy;
  let maxSeq = 0;
  values.forEach(r => {
    const rn = String(r[idx.RunningNo] || '');
    if (rn.indexOf(prefix) === 0 && rn.indexOf(suffix) === rn.length - suffix.length) {
      const mid = rn.substring(prefix.length, rn.length - suffix.length);
      const n = parseInt(mid, 10);
      if (!isNaN(n) && n > maxSeq) maxSeq = n;
    }
  });
  const next = maxSeq + 1;
  const padded = ('000' + next).slice(-3);
  return code + '/' + padded + '/' + yy;
}

// ============================================================
// APPROVAL
// ============================================================
function getApprovalView_(transferId, token) {
  const found = findTransferRow_(transferId);
  if (!found) return { ok: false, error: 'ไม่พบใบโอนย้ายนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  const items = getTransferItems_(transferId);
  const obj = Object.assign({}, found.obj);
  delete obj.ApprovalToken;
  obj.Items = items;
  return { ok: true, data: obj };
}

function decideTransfer_(body) {
  const transferId = body.transferId;
  const token = body.token;
  const decision = body.decision; // 'Approved' | 'Rejected'
  const found = findTransferRow_(transferId);
  if (!found) return { ok: false, error: 'ไม่พบใบโอนย้ายนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  if (found.obj.Status !== STATUS.PENDING) {
    return { ok: false, error: 'ใบโอนย้ายนี้ถูกดำเนินการไปแล้ว (สถานะปัจจุบัน: ' + found.obj.Status + ')' };
  }
  if (decision !== STATUS.APPROVED && decision !== STATUS.REJECTED) {
    return { ok: false, error: 'decision ไม่ถูกต้อง' };
  }

  const sh = getSS_().getSheetByName(SHEETS.TRANSFERS);
  const idx = found.idx;
  const rowNum = found.rowNum;
  sh.getRange(rowNum, idx.Status + 1).setValue(decision);
  sh.getRange(rowNum, idx.ApprovedAt + 1).setValue(new Date());
  sh.getRange(rowNum, idx.ApproverComment + 1).setValue(body.comment || '');

  logActivity_(transferId, decision.toUpperCase(), found.obj.ApproverName || found.obj.ApproverEmail, body.comment || '');

  sendDecisionNotification_(found.obj, decision, body.comment || '');

  return { ok: true, data: { transferId, status: decision } };
}

// ============================================================
// IMAGE UPLOAD (Google Drive)
// ============================================================
function uploadImage_(body) {
  try {
    const folder = getOrCreateFolder_(CONFIG.DRIVE_FOLDER_NAME);
    const contentType = body.contentType || 'image/jpeg';
    const bytes = Utilities.base64Decode(body.base64.split(',').pop());
    const blob = Utilities.newBlob(bytes, contentType, (body.fileName || 'asset') + '_' + Date.now() + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    return { ok: true, data: { url: url, fileId: file.getId() } };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

// ============================================================
// EMAIL
// ============================================================
function sendApprovalEmail_(transferId, runningNo, body, items, token) {
  try {
    const approveUrl = CONFIG.FRONTEND_URL + '?view=approve&id=' + encodeURIComponent(transferId) + '&token=' + encodeURIComponent(token);
    const itemsHtml = items.map((it, i) => (
      '<tr>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">' + (i + 1) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + (it.imageUrl ? '<img src="' + it.imageUrl + '" width="60" style="border-radius:4px;">' : '-') + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetId) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetName) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.remark || '') + '</td>' +
      '</tr>'
    )).join('');

    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:640px;margin:auto;">' +
      '<h2 style="color:#1a3c6e;">' + CONFIG.COMPANY_NAME + '</h2>' +
      '<h3>ใบโอนย้ายทรัพย์สิน เลขที่ ' + runningNo + '</h3>' +
      '<p><b>เรื่อง:</b> ' + escapeHtml_(body.subject === 'อื่นๆ' ? body.subjectOther : body.subject) + '</p>' +
      '<p><b>เพื่อ:</b> ' + escapeHtml_(body.purpose || '-') + '</p>' +
      '<p><b>จาก:</b> ' + escapeHtml_(body.fromDept) + ' &nbsp; <b>ไปยัง:</b> ' + escapeHtml_(body.toDept) + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
      '<tr style="background:#f0f4f8;"><th style="border:1px solid #ddd;padding:6px;">#</th><th style="border:1px solid #ddd;padding:6px;">รูป</th><th style="border:1px solid #ddd;padding:6px;">รหัส</th><th style="border:1px solid #ddd;padding:6px;">รายการ</th><th style="border:1px solid #ddd;padding:6px;">หมายเหตุ</th></tr>' +
      itemsHtml +
      '</table>' +
      '<p style="margin-top:20px;">กรุณาตรวจสอบรายละเอียดฉบับเต็มและพิจารณาอนุมัติที่ลิงก์ด้านล่าง:</p>' +
      '<p><a href="' + approveUrl + '" style="background:#1a3c6e;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;display:inline-block;">เปิดใบโอนย้ายเพื่อพิจารณาอนุมัติ</a></p>' +
      '<p style="color:#888;font-size:12px;">ผู้ขอ: ' + escapeHtml_(body.createdBy || '-') + '</p>' +
      '</div>';

    MailApp.sendEmail({
      to: body.approverEmail,
      subject: '[ขออนุมัติ] ใบโอนย้ายทรัพย์สิน ' + runningNo + ' — ' + CONFIG.COMPANY_NAME,
      htmlBody: html
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function sendDecisionNotification_(transferObj, decision, comment) {
  try {
    if (!transferObj.CreatedByEmail) return;
    const statusThai = decision === STATUS.APPROVED ? 'อนุมัติ' : 'ไม่อนุมัติ';
    const color = decision === STATUS.APPROVED ? '#1a7d3c' : '#c0392b';
    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:auto;">' +
      '<h3>ใบโอนย้ายทรัพย์สิน เลขที่ ' + transferObj.RunningNo + '</h3>' +
      '<p style="font-size:16px;">สถานะ: <b style="color:' + color + ';">' + statusThai + '</b></p>' +
      (comment ? '<p><b>ความเห็นผู้อนุมัติ:</b> ' + escapeHtml_(comment) + '</p>' : '') +
      '<p>โดย: ' + escapeHtml_(transferObj.ApproverName || transferObj.ApproverEmail) + '</p>' +
      '</div>';
    MailApp.sendEmail({
      to: transferObj.CreatedByEmail,
      subject: '[' + statusThai + '] ใบโอนย้ายทรัพย์สิน ' + transferObj.RunningNo,
      htmlBody: html
    });
  } catch (err) {
    Logger.log('sendDecisionNotification_ error: ' + err);
  }
}

// ============================================================
// LOG
// ============================================================
function logActivity_(transferId, action, by, detail) {
  const sh = getSS_().getSheetByName(SHEETS.LOG);
  sh.appendRow([new Date(), transferId, action, by, detail]);
}

// ============================================================
// UTIL
// ============================================================
function indexMap_(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  return idx;
}

function rowToObj_(row, idx) {
  const obj = {};
  Object.keys(idx).forEach(key => {
    let v = row[idx[key]];
    if (v instanceof Date) v = v.toISOString();
    obj[key] = v;
  });
  return obj;
}

function escapeHtml_(str) {
  return String(str || '').replace(/[&<>"']/g, function (m) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
  });
}
