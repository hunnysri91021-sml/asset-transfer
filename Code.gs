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
 *  7. เปิดชีต Users แล้วเพิ่มผู้ใช้ระดับ Admin อย่างน้อย 1 คน (Username, Password, Role=admin) ด้วยตนเอง
 *       ก่อนเริ่มใช้งานจริง — ไม่เช่นนั้นจะไม่มีใคร login เข้าระบบได้เลย (ไม่มีรหัสผ่านสำรองอีกต่อไป)
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
  USERS: 'Users',
  TRANSFER_QUEUE: 'TransferQueue',
  TRANSFERS: 'Transfers',
  ITEMS: 'TransferItems',
  SALES: 'Sales',
  SALE_ITEMS: 'SaleItems',
  WRITEOFFS: 'WriteOffs',
  WRITEOFF_ITEMS: 'WriteOffItems',
  LOG: 'ActivityLog'
};

const HEADERS = {
  ASSETS: ['AssetID', 'AssetName', 'Department', 'Division', 'WorkGroup', 'PurchaseDate', 'PurchasePrice', 'BookValue', 'Custodian', 'Location', 'Tag', 'ScrapPrice', 'MinSalePrice', 'ImageURL', 'ImageURLOverride', 'UpdatedAt', 'SyncFlag', 'SyncNote'],
  DEPT_CODES: ['DeptName', 'Code', 'ApproverName', 'ApproverEmail', 'SkipApprovalEmail', 'StartSeqTransfer', 'StartSeqSale', 'StartSeqWriteOff'],
  USERS: ['Username', 'Password', 'Role', 'Departments', 'CreatedAt'],
  TRANSFER_QUEUE: ['AssetID', 'Purpose', 'AddedBy', 'AddedAt'],
  TRANSFERS: ['TransferID', 'RunningNo', 'CreatedAt', 'Subject', 'SubjectOther', 'Purpose', 'FromDept', 'FromDeptCode', 'ToDept', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail'],
  ITEMS: ['TransferID', 'LineNo', 'AssetID', 'AssetName', 'FromDeptName', 'FromSignName', 'ToDeptName', 'ToSignName', 'Remark', 'ImageURL'],
  SALES: ['SaleID', 'RunningNo', 'CreatedAt', 'FromDept', 'FromDeptCode', 'Buyer', 'Remark', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail'],
  SALE_ITEMS: ['SaleID', 'LineNo', 'AssetID', 'AssetName', 'ScrapPrice', 'AuctionPrice', 'SalePrice', 'Remark', 'ImageURL'],
  WRITEOFFS: ['WriteOffID', 'RunningNo', 'CreatedAt', 'FromDept', 'FromDeptCode', 'Reason', 'Remark', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail'],
  WRITEOFF_ITEMS: ['WriteOffID', 'LineNo', 'AssetID', 'AssetName', 'ScrapPrice', 'Remark', 'ImageURL'],
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
  REJECTED: 'Rejected',
  VOIDED: 'Voided' // ใช้เมื่อ Admin กด "คืนสถานะใช้งาน" ยกเลิกผลของใบขาย/ใบตัดชำรุดที่อนุมัติแล้ว
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
  ensureSheet_(ss, SHEETS.USERS, HEADERS.USERS);
  ensureSheet_(ss, SHEETS.TRANSFER_QUEUE, HEADERS.TRANSFER_QUEUE);
  ensureSheet_(ss, SHEETS.TRANSFERS, HEADERS.TRANSFERS);
  ensureSheet_(ss, SHEETS.ITEMS, HEADERS.ITEMS);
  ensureSheet_(ss, SHEETS.SALES, HEADERS.SALES);
  ensureSheet_(ss, SHEETS.SALE_ITEMS, HEADERS.SALE_ITEMS);
  ensureSheet_(ss, SHEETS.WRITEOFFS, HEADERS.WRITEOFFS);
  ensureSheet_(ss, SHEETS.WRITEOFF_ITEMS, HEADERS.WRITEOFF_ITEMS);
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
      case 'getAssetsFull':
        result = { ok: true, data: getAssetsFull_(e.parameter.q || '') };
        break;
      case 'getDeptCodes':
        result = { ok: true, data: getDeptCodes_() };
        break;
      case 'getTransferQueue':
        result = { ok: true, data: getTransferQueue_() };
        break;
      case 'getSaleQueue':
        result = { ok: true, data: getSaleQueue_() };
        break;
      case 'getWriteOffQueue':
        result = { ok: true, data: getWriteOffQueue_() };
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
      case 'getSales':
        result = { ok: true, data: getSales_(e.parameter) };
        break;
      case 'getSale':
        result = { ok: true, data: getSaleFull_(e.parameter.id) };
        break;
      case 'getSaleApprovalView':
        result = getSaleApprovalView_(e.parameter.id, e.parameter.token);
        break;
      case 'getWriteOffs':
        result = { ok: true, data: getWriteOffs_(e.parameter) };
        break;
      case 'getWriteOff':
        result = { ok: true, data: getWriteOffFull_(e.parameter.id) };
        break;
      case 'getWriteOffApprovalView':
        result = getWriteOffApprovalView_(e.parameter.id, e.parameter.token);
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
      case 'addToTransferQueue':
        result = addToTransferQueue_(body);
        break;
      case 'removeFromTransferQueue':
        result = removeFromTransferQueue_(body);
        break;
      case 'addToSaleQueue':
        result = addToSaleQueue_(body);
        break;
      case 'removeFromSaleQueue':
        result = removeFromSaleQueue_(body);
        break;
      case 'addToWriteOffQueue':
        result = addToWriteOffQueue_(body);
        break;
      case 'removeFromWriteOffQueue':
        result = removeFromWriteOffQueue_(body);
        break;
      case 'createTransfer':
        result = createTransfer_(body);
        break;
      case 'decideTransfer':
        result = decideTransfer_(body);
        break;
      case 'createSale':
        result = createSale_(body);
        break;
      case 'decideSale':
        result = decideSale_(body);
        break;
      case 'createWriteOff':
        result = createWriteOff_(body);
        break;
      case 'decideWriteOff':
        result = decideWriteOff_(body);
        break;
      case 'uploadImage':
        result = uploadImage_(body);
        break;
      case 'updateAssetImage':
        result = updateAssetImage_(body);
        break;
      case 'login':
        result = login_(body);
        break;
      case 'adminGetUsers':
        result = getUsers_(body);
        break;
      case 'adminSaveUser':
        result = adminSaveUser_(body);
        break;
      case 'adminDeleteUser':
        result = adminDeleteUser_(body);
        break;
      case 'adminSaveAsset':
        result = adminSaveAsset_(body);
        break;
      case 'adminDeleteAsset':
        result = adminDeleteAsset_(body);
        break;
      case 'adminGetSettings':
        result = adminGetSettings_(body);
        break;
      case 'adminSaveSourceSheetLink':
        result = adminSaveSourceSheetLink_(body);
        break;
      case 'adminSaveScrapRate':
        result = adminSaveScrapRate_(body);
        break;
      case 'adminBackfillAssetTags':
        result = adminBackfillAssetTags_(body);
        break;
      case 'adminSyncFromSource':
        result = adminSyncFromSource_(body);
        break;
      case 'adminRestoreAsset':
        result = adminRestoreAsset_(body);
        break;
      case 'adminSaveDept':
        result = adminSaveDept_(body);
        break;
      case 'adminDeleteDept':
        result = adminDeleteDept_(body);
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
  const disposed = getDisposedAssetStatus_();
  let rows = getAssetsRaw_(q);
  rows = rows.filter(r => !disposed[String(r.AssetID)]);
  return rows;
}

// รายการทรัพย์สินทั้งหมด (ไม่ซ่อนรายการที่ขาย/ตัดชำรุดแล้ว) พร้อมสถานะปัจจุบัน — ใช้กับแถบ "รายการทรัพย์สิน" และ "ข้อมูลรวม"
function getAssetsFull_(q) {
  const disposed = getDisposedAssetStatus_();
  const rows = getAssetsRaw_(q);
  rows.forEach(r => { r.AssetStatus = disposed[String(r.AssetID)] || 'Active'; });
  return rows;
}

function getAssetsRaw_(q) {
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  let rows = values.map(r => rowToObj_(r, idx));
  if (q) {
    const qq = q.toString().toLowerCase();
    rows = rows.filter(r =>
      String(r.AssetID).toLowerCase().indexOf(qq) !== -1 ||
      String(r.AssetName).toLowerCase().indexOf(qq) !== -1 ||
      String(r.Custodian).toLowerCase().indexOf(qq) !== -1
    );
  }
  // resolve display image: override wins over original
  rows.forEach(r => { r.DisplayImage = r.ImageURLOverride || r.ImageURL || ''; });
  // ราคาซาก คำนวณอัตโนมัติตามหลักบัญชี = ราคาซื้อ x เปอร์เซ็นต์ที่ตั้งค่าไว้ (ไม่รับค่าที่พิมพ์เอง/ซิงค์จากภายนอกอีกต่อไป)
  const scrapRate = getScrapRatePercent_();
  rows.forEach(r => {
    const purchasePrice = parseFloat(r.PurchasePrice) || 0;
    r.ScrapPrice = Math.round(purchasePrice * scrapRate / 100 * 100) / 100;
  });
  return rows;
}

// ============================================================
// ASSET QUEUE — ทรัพย์สินที่มาร์คไว้ล่วงหน้าจากหน้า "รายการทรัพย์สิน" เพื่อรอออกใบโอนย้าย/ขายออก/ตัดชำรุด
// ใช้ชีตเดียวกันร่วมกันทั้ง 3 ประเภท แยกด้วยคอลัมน์ Purpose
// ============================================================
const QUEUE_PURPOSES = { TRANSFER: 'Transfer', SALE: 'Sale', WRITEOFF: 'WriteOff' };

function getAssetQueue_(purpose) {
  const sh = getSS_().getSheetByName(SHEETS.TRANSFER_QUEUE);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  const queueRows = values
    .filter(r => r[idx.AssetID])
    .map(r => rowToObj_(r, idx))
    .filter(r => (r.Purpose || QUEUE_PURPOSES.TRANSFER) === purpose);

  const disposed = getDisposedAssetStatus_();
  const assetsById = {};
  getAssetsRaw_().forEach(a => { assetsById[String(a.AssetID)] = a; });

  return queueRows.map(q => {
    const a = assetsById[String(q.AssetID)] || {};
    return {
      AssetID: q.AssetID,
      AddedBy: q.AddedBy,
      AddedAt: q.AddedAt,
      AssetName: a.AssetName || '',
      Department: a.Department || '',
      DisplayImage: a.DisplayImage || '',
      PurchasePrice: a.PurchasePrice || 0,
      BookValue: a.BookValue || 0,
      ScrapPrice: a.ScrapPrice || 0,
      MinSalePrice: a.MinSalePrice || 0,
      AssetStatus: disposed[String(q.AssetID)] || 'Active'
    };
  }).sort((x, y) => new Date(x.AddedAt) - new Date(y.AddedAt));
}

function addToAssetQueue_(body, purpose) {
  const assetId = String(body.assetId || '').trim();
  if (!assetId) return { ok: false, error: 'กรุณาระบุรหัสทรัพย์สิน' };

  const disposed = getDisposedAssetStatus_();
  if (disposed[assetId]) return { ok: false, error: 'ทรัพย์สินนี้ถูกขาย/ตัดชำรุดไปแล้ว ไม่สามารถเพิ่มเข้าคิวได้' };

  // ล็อกช่วงตรวจสอบรายการซ้ำ + เพิ่มแถว กันซ้ำเมื่อมีคนกดเพิ่มทรัพย์สินเดียวกันเข้าคิวพร้อมกัน
  return withLock_(() => {
    const sh = getSS_().getSheetByName(SHEETS.TRANSFER_QUEUE);
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const idx = indexMap_(headers);
    // ชีตนี้เดิมมีแค่ AssetID/AddedBy/AddedAt (ก่อนเพิ่มคิวขาย/ตัดชำรุด) คอลัมน์ Purpose อาจยังไม่ถูกสร้างถ้ายังไม่ได้รัน setup() ใหม่
    if (idx.Purpose === undefined && purpose !== QUEUE_PURPOSES.TRANSFER) {
      return { ok: false, error: 'ชีต TransferQueue ยังไม่มีคอลัมน์ Purpose กรุณาให้ Admin รันฟังก์ชัน setup() ใหม่ใน Apps Script ก่อนใช้งานคิวขาย/คิวตัดชำรุด' };
    }
    for (let i = 1; i < values.length; i++) {
      const rowPurpose = idx.Purpose !== undefined ? (values[i][idx.Purpose] || QUEUE_PURPOSES.TRANSFER) : QUEUE_PURPOSES.TRANSFER;
      if (String(values[i][idx.AssetID]) === assetId && rowPurpose === purpose) return { ok: true, data: { alreadyQueued: true } };
    }
    // เขียนตามตำแหน่งคอลัมน์จริงในชีต (idx) แทนการอิงลำดับคงที่ กันค่าคลาดเคลื่อนคอลัมน์เหมือนที่เคยเกิดกับชีต Users
    const newRow = headers.map(() => '');
    newRow[idx.AssetID] = assetId;
    if (idx.Purpose !== undefined) newRow[idx.Purpose] = purpose;
    newRow[idx.AddedBy] = body.addedBy || '';
    newRow[idx.AddedAt] = new Date();
    sh.appendRow(newRow);
    return { ok: true, data: { alreadyQueued: false } };
  });
}

function removeFromAssetQueue_(body, purpose) {
  const assetIds = (body.assetIds || []).map(String);
  if (!assetIds.length) return { ok: false, error: 'กรุณาระบุรายการที่ต้องการลบออกจากคิว' };
  const sh = getSS_().getSheetByName(SHEETS.TRANSFER_QUEUE);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  let removed = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const rowPurpose = values[i][idx.Purpose] || QUEUE_PURPOSES.TRANSFER;
    if (rowPurpose === purpose && assetIds.indexOf(String(values[i][idx.AssetID])) !== -1) {
      sh.deleteRow(i + 1);
      removed++;
    }
  }
  return { ok: true, data: { removed } };
}

function getTransferQueue_() { return getAssetQueue_(QUEUE_PURPOSES.TRANSFER); }
function addToTransferQueue_(body) { return addToAssetQueue_(body, QUEUE_PURPOSES.TRANSFER); }
function removeFromTransferQueue_(body) { return removeFromAssetQueue_(body, QUEUE_PURPOSES.TRANSFER); }

function getSaleQueue_() { return getAssetQueue_(QUEUE_PURPOSES.SALE); }
function addToSaleQueue_(body) { return addToAssetQueue_(body, QUEUE_PURPOSES.SALE); }
function removeFromSaleQueue_(body) { return removeFromAssetQueue_(body, QUEUE_PURPOSES.SALE); }

function getWriteOffQueue_() { return getAssetQueue_(QUEUE_PURPOSES.WRITEOFF); }
function addToWriteOffQueue_(body) { return addToAssetQueue_(body, QUEUE_PURPOSES.WRITEOFF); }
function removeFromWriteOffQueue_(body) { return removeFromAssetQueue_(body, QUEUE_PURPOSES.WRITEOFF); }

// เมื่อทรัพย์สินถูกขาย/ตัดชำรุดจนอนุมัติเสร็จสมบูรณ์แล้ว (ไม่ว่างอีกต่อไป) ให้เอาออกจากคิวรอทุกประเภท
// กันไม่ให้ค้างเป็นรายการ "ผี" ในคิวโอนย้าย/ขาย/ตัดชำรุดที่เลือกไว้ก่อนหน้า
function purgeAssetFromAllQueues_(assetIds) {
  if (!assetIds || !assetIds.length) return;
  removeFromAssetQueue_({ assetIds: assetIds }, QUEUE_PURPOSES.TRANSFER);
  removeFromAssetQueue_({ assetIds: assetIds }, QUEUE_PURPOSES.SALE);
  removeFromAssetQueue_({ assetIds: assetIds }, QUEUE_PURPOSES.WRITEOFF);
}

// AssetStatus (Active/Sold/WrittenOff) คำนวณสดจากเอกสารขาย/ตัดชำรุดที่อนุมัติแล้วเสมอ (ดู getDisposedAssetStatus_)
// แต่ค่านั้นไม่เคยถูกเขียนกลับไปที่ชีต Assets เอง ทำให้เปิดชีตดูตรงๆ หรือดูสรุปแยกตามแท็กสถานะในหน้า Dashboard
// แล้วไม่เห็นความเปลี่ยนแปลง — ฟังก์ชันนี้เขียนคอลัมน์ Tag ในชีต Assets ให้ตรงกับผลจริงด้วย ทุกครั้งที่มีการ
// อนุมัติขาย/ตัดชำรุด (หรือคืนสถานะใช้งาน) เพื่อให้ทั้งชีตดิบและ Dashboard สอดคล้องกัน
function setAssetsTag_(assetIds, tagValue) {
  if (!assetIds || !assetIds.length) return;
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  if (idx.Tag === undefined || idx.AssetID === undefined) return;
  const idSet = {};
  assetIds.forEach(id => { idSet[String(id)] = true; });
  for (let i = 1; i < values.length; i++) {
    if (idSet[String(values[i][idx.AssetID])]) {
      sh.getRange(i + 1, idx.Tag + 1).setValue(tagValue);
    }
  }
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

// ทรัพย์สินที่มีใบขายออก หรือใบตัดชำรุด ซึ่งอนุมัติแล้ว ถือว่าสิ้นสภาพการใช้งานจริง จึงซ่อนจากรายการหลัก
// คืนค่า map ของ AssetID -> 'Sold' | 'WrittenOff' (เฉพาะรายการที่สิ้นสภาพแล้ว)
function getDisposedAssetStatus_() {
  const status = {};
  markDisposedFromDocs_(status, SHEETS.SALES, SHEETS.SALE_ITEMS, 'SaleID', 'Sold');
  markDisposedFromDocs_(status, SHEETS.WRITEOFFS, SHEETS.WRITEOFF_ITEMS, 'WriteOffID', 'WrittenOff');
  return status;
}

function markDisposedFromDocs_(status, docSheetName, itemSheetName, docIdField, label) {
  const docSh = getSS_().getSheetByName(docSheetName);
  const docValues = docSh.getDataRange().getValues();
  const docHeaders = docValues.shift();
  const docIdx = indexMap_(docHeaders);
  const approvedDocIds = {};
  docValues.forEach(r => {
    if (r[docIdx.Status] === STATUS.APPROVED) approvedDocIds[String(r[docIdx[docIdField]])] = true;
  });

  const itemSh = getSS_().getSheetByName(itemSheetName);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  itemValues.forEach(r => {
    if (approvedDocIds[String(r[itemIdx[docIdField]])]) status[String(r[itemIdx.AssetID])] = label;
  });
}

// ============================================================
// ADMIN — แก้ไขข้อมูลทรัพย์สินหลัก (ชีต Assets ที่ทีมบัญชี upload เข้ามา)
// ============================================================
// รหัสผ่าน admin ต้องตรงกับผู้ใช้ที่มีสิทธิ์ admin ในชีต Users เท่านั้น (ไม่มีรหัสผ่านกลางสำรองอีกต่อไป)
function checkAdminPassword_(pw) {
  const p = String(pw || '');
  if (!p) return false;
  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.Password]) === p && String(values[i][idx.Role]) === 'admin') return true;
  }
  return false;
}

// เข้าสู่ระบบด้วยชื่อผู้ใช้ + รหัสผ่านที่ Admin ตั้งไว้ในชีต Users เท่านั้น
function login_(body) {
  const username = String(body.username || '').trim();
  const password = String(body.password || '');
  if (!username || !password) return { ok: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };

  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.Username]) === username && String(values[i][idx.Password]) === password) {
      return { ok: true, data: {
        username,
        role: values[i][idx.Role] === 'admin' ? 'admin' : 'user',
        departments: parseDepartments_(values[i][idx.Departments])
      } };
    }
  }
  return { ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function parseDepartments_(v) {
  return String(v || '').split(',').map(s => s.trim()).filter(Boolean);
}

// ค้นหาผู้ใช้จากรหัสผ่านที่ส่งมา (ไม่ต้องรู้ username ล่วงหน้า) เพื่อตรวจสอบสิทธิ์ role/หน่วยงานที่แก้ไขได้
function getRequestingUser_(pw) {
  const p = String(pw || '');
  if (!p) return null;
  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.Password]) === p) {
      return {
        username: values[i][idx.Username],
        role: String(values[i][idx.Role]) === 'admin' ? 'admin' : 'user',
        departments: parseDepartments_(values[i][idx.Departments])
      };
    }
  }
  return null;
}

// Admin แก้ไข/เพิ่ม/ลบได้ทุกหน่วยงาน ส่วน User แก้ไขได้เฉพาะหน่วยงานที่ Admin กำหนดสิทธิ์ให้เท่านั้น
function canManageDept_(user, dept) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return user.departments.indexOf(String(dept || '').trim()) !== -1;
}

function getUsers_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  const users = values.filter(r => r[idx.Username]).map(r => ({
    Username: r[idx.Username],
    Role: r[idx.Role],
    Departments: parseDepartments_(r[idx.Departments]),
    CreatedAt: r[idx.CreatedAt] instanceof Date ? r[idx.CreatedAt].toISOString() : r[idx.CreatedAt]
  }));
  return { ok: true, data: users };
}

function adminSaveUser_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const u = body.user || {};
  const username = String(u.Username || '').trim();
  if (!username) return { ok: false, error: 'กรุณาระบุชื่อผู้ใช้' };
  const role = u.Role === 'admin' ? 'admin' : 'user';
  const newPassword = String(u.Password || '').trim();
  const departments = Array.isArray(u.Departments) ? u.Departments.map(d => String(d).trim()).filter(Boolean).join(',') : '';

  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  let rowNum = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.Username]) === username) { rowNum = i + 1; break; }
  }

  if (rowNum === -1) {
    if (!newPassword) return { ok: false, error: 'กรุณาระบุรหัสผ่านสำหรับผู้ใช้ใหม่' };
    // เขียนตามตำแหน่งคอลัมน์จริงในชีต (idx) แทนการอิงลำดับคงที่ เพื่อไม่ให้ค่าคลาดเคลื่อนคอลัมน์
    // ถ้าชีตยังไม่มีคอลัมน์ Departments (ยังไม่ได้รัน setup() ใหม่)
    const newRow = headers.map(() => '');
    newRow[idx.Username] = username;
    newRow[idx.Password] = newPassword;
    newRow[idx.Role] = role;
    if (idx.Departments !== undefined) newRow[idx.Departments] = departments;
    newRow[idx.CreatedAt] = new Date();
    sh.appendRow(newRow);
    logActivity_('', 'ADMIN_SAVE_USER', 'admin', 'เพิ่มผู้ใช้ ' + username);
    return { ok: true, data: { created: true } };
  }
  sh.getRange(rowNum, idx.Role + 1).setValue(role);
  if (idx.Departments !== undefined) sh.getRange(rowNum, idx.Departments + 1).setValue(departments);
  if (newPassword) sh.getRange(rowNum, idx.Password + 1).setValue(newPassword);
  logActivity_('', 'ADMIN_SAVE_USER', 'admin', 'แก้ไขผู้ใช้ ' + username);
  return { ok: true, data: { created: false } };
}

function adminDeleteUser_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const username = String(body.username || '').trim();
  if (!username) return { ok: false, error: 'กรุณาระบุชื่อผู้ใช้' };
  const sh = getSS_().getSheetByName(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.Username]) === username) {
      sh.deleteRow(i + 1);
      logActivity_('', 'ADMIN_DELETE_USER', 'admin', 'ลบผู้ใช้ ' + username);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบผู้ใช้นี้' };
}

// ScrapPrice ไม่อยู่ในนี้แล้ว เพราะคำนวณอัตโนมัติจาก PurchasePrice x เปอร์เซ็นต์ราคาซาก (ดู getScrapRatePercent_)
const ADMIN_ASSET_EDITABLE_FIELDS = ['AssetName', 'Department', 'Division', 'WorkGroup', 'PurchaseDate', 'PurchasePrice', 'BookValue', 'Custodian', 'Location', 'Tag', 'MinSalePrice', 'ImageURL'];

// แท็กสถานะการใช้งานที่ Admin ปรับได้อิสระ ไม่ต้องขออนุมัติ (แยกจากสถานะที่คำนวณจากใบขาย/ใบตัดชำรุดที่อนุมัติแล้ว)
const ASSET_TAGS = ['ใช้งาน', 'ชำรุด', 'ขาย', 'เก็บไว้ใช้', 'รอเปลี่ยนอะไหล่'];

function adminSaveAsset_(body) {
  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
  const asset = body.asset || {};
  const assetId = String(asset.AssetID || '').trim();
  if (!assetId) return { ok: false, error: 'กรุณาระบุรหัสทรัพย์สิน (AssetID)' };

  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  let rowNum = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.AssetID]) === assetId) { rowNum = i + 1; break; }
  }

  if (rowNum === -1) {
    if (!canManageDept_(user, asset.Department)) return { ok: false, error: 'ไม่มีสิทธิ์เพิ่มทรัพย์สินของหน่วยงานนี้' };
    // เขียนตามตำแหน่งคอลัมน์จริงในชีต (headers/idx) แทนการอิงลำดับคงที่ใน HEADERS.ASSETS
    // เพื่อไม่ให้ค่าคลาดเคลื่อนคอลัมน์ถ้าชีตจริงมีลำดับต่างจากค่าคงที่ (เช่น ยังไม่ได้รัน setup() ใหม่)
    const newRow = headers.map(h => {
      if (h === 'AssetID') return assetId;
      if (h === 'UpdatedAt') return new Date();
      if (ADMIN_ASSET_EDITABLE_FIELDS.indexOf(h) !== -1) return asset[h] || '';
      return '';
    });
    sh.appendRow(newRow);
    logActivity_('', 'ADMIN_ASSET_CREATE', user.username, 'เพิ่มทรัพย์สิน ' + assetId);
    return { ok: true, data: { created: true } };
  }

  const existingDept = values[rowNum - 1][idx.Department];
  if (!canManageDept_(user, existingDept)) return { ok: false, error: 'ไม่มีสิทธิ์แก้ไขทรัพย์สินของหน่วยงานนี้' };
  if (asset.Department !== undefined && String(asset.Department).trim() !== String(existingDept || '').trim() && !canManageDept_(user, asset.Department)) {
    return { ok: false, error: 'ไม่มีสิทธิ์ย้ายทรัพย์สินไปยังหน่วยงานนี้' };
  }

  ADMIN_ASSET_EDITABLE_FIELDS.forEach(f => {
    if (asset[f] !== undefined) sh.getRange(rowNum, idx[f] + 1).setValue(asset[f]);
  });
  sh.getRange(rowNum, idx.UpdatedAt + 1).setValue(new Date());
  // มีคนแก้ไข/บันทึกรายการนี้แล้ว ถือว่ารับทราบ จึงล้างป้าย "ข้อมูลใหม่จากบัญชี" / "ไม่ตรงกับบัญชี" ทิ้ง
  // (เช็ค idx !== undefined เผื่อยังไม่ได้รัน setup() ใหม่เพื่อเพิ่มคอลัมน์ SyncFlag/SyncNote ในชีต)
  if (idx.SyncFlag !== undefined) sh.getRange(rowNum, idx.SyncFlag + 1).setValue('');
  if (idx.SyncNote !== undefined) sh.getRange(rowNum, idx.SyncNote + 1).setValue('');
  logActivity_('', 'ADMIN_ASSET_UPDATE', user.username, 'แก้ไขทรัพย์สิน ' + assetId);
  return { ok: true, data: { created: false } };
}

function adminDeleteAsset_(body) {
  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
  const assetId = String(body.assetId || '').trim();
  if (!assetId) return { ok: false, error: 'กรุณาระบุรหัสทรัพย์สิน' };

  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.AssetID]) === assetId) {
      if (!canManageDept_(user, values[i][idx.Department])) return { ok: false, error: 'ไม่มีสิทธิ์ลบทรัพย์สินของหน่วยงานนี้' };
      sh.deleteRow(i + 1);
      logActivity_('', 'ADMIN_ASSET_DELETE', user.username, 'ลบทรัพย์สิน ' + assetId);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบทรัพย์สิน: ' + assetId };
}

// คืนสถานะทรัพย์สินที่เคยขาย/ตัดชำรุด (อนุมัติแล้ว) กลับมาเป็น "ใช้งาน" — โดยยกเลิก (Voided)
// ใบขาย/ใบตัดชำรุดที่เกี่ยวข้องทั้งหมด แทนการลบประวัติ เพื่อให้ยังตรวจสอบย้อนหลังได้
function adminRestoreAsset_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const assetId = String(body.assetId || '').trim();
  if (!assetId) return { ok: false, error: 'กรุณาระบุรหัสทรัพย์สิน' };

  const voidedCount =
    voidApprovedDocsForAsset_(SHEETS.SALES, SHEETS.SALE_ITEMS, 'SaleID', assetId) +
    voidApprovedDocsForAsset_(SHEETS.WRITEOFFS, SHEETS.WRITEOFF_ITEMS, 'WriteOffID', assetId);

  if (!voidedCount) return { ok: false, error: 'ไม่พบใบขาย/ใบตัดชำรุดที่อนุมัติแล้วของทรัพย์สินนี้' };

  setAssetsTag_([assetId], 'ใช้งาน');
  logActivity_('', 'ADMIN_RESTORE_ASSET', 'admin', 'คืนสถานะใช้งานทรัพย์สิน ' + assetId);
  return { ok: true, data: { voidedCount } };
}

function voidApprovedDocsForAsset_(docSheetName, itemSheetName, docIdField, assetId) {
  const itemSh = getSS_().getSheetByName(itemSheetName);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  const docIds = {};
  itemValues.forEach(r => {
    if (String(r[itemIdx.AssetID]) === assetId) docIds[String(r[itemIdx[docIdField]])] = true;
  });
  if (!Object.keys(docIds).length) return 0;

  const docSh = getSS_().getSheetByName(docSheetName);
  const docValues = docSh.getDataRange().getValues();
  const docHeaders = docValues[0];
  const docIdx = indexMap_(docHeaders);
  let count = 0;
  for (let i = 1; i < docValues.length; i++) {
    const id = String(docValues[i][docIdx[docIdField]]);
    if (docIds[id] && docValues[i][docIdx.Status] === STATUS.APPROVED) {
      docSh.getRange(i + 1, docIdx.Status + 1).setValue(STATUS.VOIDED);
      count++;
    }
  }
  return count;
}

// ============================================================
// ADMIN — เชื่อมข้อมูลจาก Google Sheet ต้นทาง (เช่น AppSheet.ViewData ของทีมบัญชี)
// ============================================================
// ทีมบัญชีอัปโหลดแท็บใหม่ทุกครั้งโดยตั้งชื่อแบบ "AppSheet.ViewData.<วันที่>" (เช่น AppSheet.ViewData.2026-08-13)
// จึงจับคู่ด้วย prefix แล้วเลือกแท็บที่มีชื่อ (วันที่) ล่าสุดโดยอัตโนมัติ แทนชื่อคงที่
const SOURCE_SHEET_TAB_PREFIX = 'AppSheet.ViewData';
const SOURCE_SHEET_URL_PROP = 'SOURCE_SHEET_URL';
const SOURCE_SYNC_FIELDS = ['AssetName', 'Department', 'Division', 'WorkGroup', 'PurchaseDate', 'PurchasePrice', 'BookValue', 'Custodian', 'Location', 'MinSalePrice', 'ImageURL'];

function getSourceSheetUrl_() {
  return PropertiesService.getScriptProperties().getProperty(SOURCE_SHEET_URL_PROP) || '';
}

// เปอร์เซ็นต์ราคาซาก (ตามหลักบัญชี) ที่ใช้คำนวณ ราคาซาก = ราคาซื้อ x เปอร์เซ็นต์นี้ ค่าเริ่มต้น 5%
const SCRAP_RATE_PROP = 'SCRAP_RATE_PERCENT';
const DEFAULT_SCRAP_RATE_PERCENT = 5;
function getScrapRatePercent_() {
  const v = PropertiesService.getScriptProperties().getProperty(SCRAP_RATE_PROP);
  const n = parseFloat(v);
  return isNaN(n) ? DEFAULT_SCRAP_RATE_PERCENT : n;
}

function adminSaveScrapRate_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const rate = parseFloat(body.scrapRatePercent);
  if (isNaN(rate) || rate < 0 || rate > 100) return { ok: false, error: 'กรุณาระบุเปอร์เซ็นต์ราคาซากระหว่าง 0-100' };
  PropertiesService.getScriptProperties().setProperty(SCRAP_RATE_PROP, String(rate));
  logActivity_('', 'ADMIN_SET_SCRAP_RATE', 'admin', 'ตั้งค่าเปอร์เซ็นต์ราคาซาก ' + rate + '%');
  return { ok: true };
}

// setAssetsTag_ เขียนแท็ก "ขาย"/"ชำรุด" ให้อัตโนมัติเฉพาะตอนใบขาย/ตัดชำรุดอนุมัติใหม่นับจากตอนที่เพิ่มฟีเจอร์นี้
// (ดู createSale_/createWriteOff_/decideSale_/decideWriteOff_) — รายการที่ขาย/ตัดชำรุดไปแล้ว "ก่อนหน้านั้น" ไม่ถูกไล่ย้อนหลังให้
// ฟังก์ชันนี้ให้ Admin กดครั้งเดียวเพื่อไล่เช็คสถานะจริง (ตามเอกสารที่อนุมัติแล้วทั้งหมด) แล้วปรับ Tag ในชีต Assets ให้ตรงกันย้อนหลัง
function adminBackfillAssetTags_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const disposed = getDisposedAssetStatus_();
  const soldIds = [];
  const writtenOffIds = [];
  Object.keys(disposed).forEach(assetId => {
    if (disposed[assetId] === 'Sold') soldIds.push(assetId);
    else if (disposed[assetId] === 'WrittenOff') writtenOffIds.push(assetId);
  });
  setAssetsTag_(soldIds, 'ขาย');
  setAssetsTag_(writtenOffIds, 'ชำรุด');
  logActivity_('', 'ADMIN_BACKFILL_TAGS', 'admin', 'ซิงค์แท็กสถานะย้อนหลัง: ขาย ' + soldIds.length + ' รายการ ตัดชำรุด ' + writtenOffIds.length + ' รายการ');
  return { ok: true, data: { soldCount: soldIds.length, writtenOffCount: writtenOffIds.length } };
}

function findSourceSheet_(srcSs) {
  const sheets = srcSs.getSheets();
  const matches = sheets.filter(function (s) { return s.getName().indexOf(SOURCE_SHEET_TAB_PREFIX) === 0; });
  if (matches.length) {
    matches.sort(function (a, b) { return b.getName().localeCompare(a.getName()); });
    return matches[0];
  }
  return srcSs.getSheetByName('ViewData') || sheets[0];
}

function adminGetSettings_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  return { ok: true, data: { sourceSheetUrl: getSourceSheetUrl_(), scrapRatePercent: getScrapRatePercent_() } };
}

function adminSaveSourceSheetLink_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const url = String(body.url || '').trim();
  if (url && !/^https:\/\/docs\.google\.com\/spreadsheets\//.test(url)) {
    return { ok: false, error: 'กรุณาระบุลิงก์ Google Sheets ที่ถูกต้อง (https://docs.google.com/spreadsheets/...)' };
  }
  PropertiesService.getScriptProperties().setProperty(SOURCE_SHEET_URL_PROP, url);
  logActivity_('', 'ADMIN_SET_SOURCE_LINK', 'admin', 'ตั้งค่าลิงก์ข้อมูลต้นทาง');
  return { ok: true };
}

// ดึงข้อมูลจากชีตต้นทาง จับคู่คอลัมน์ตามชื่อหัวตาราง: อัปเดตแถวที่มี AssetID ตรงกับที่มีอยู่แล้ว
// และเพิ่มแถวใหม่ให้กับ AssetID ที่ยังไม่มีในระบบ (ไม่ลบแถวเดิมที่ไม่พบในต้นทาง เพื่อไม่ทับรายการที่ Admin เพิ่ม/แก้ไขเองในระบบ)
// ป้ายกำกับฟิลด์ (ภาษาไทย) สำหรับสรุปรายละเอียดความไม่ตรงกันให้ผู้ใช้อ่านง่าย
const SYNC_FIELD_LABELS_TH = {
  AssetName: 'ชื่อทรัพย์สิน', Department: 'หน่วยงาน', Division: 'ฝ่าย', WorkGroup: 'กลุ่มงาน',
  PurchaseDate: 'วันที่ซื้อ', PurchasePrice: 'ราคาซื้อ', BookValue: 'มูลค่าตามบัญชี',
  Custodian: 'ผู้ดูแล', Location: 'สถานที่', MinSalePrice: 'ราคาขายขั้นต่ำ', ImageURL: 'รูปภาพ'
};

function normalizeCompareValue_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof v === 'number') return String(Math.round(v * 100) / 100);
  return String(v == null ? '' : v).trim();
}

function valuesDiffer_(a, b) {
  return normalizeCompareValue_(a) !== normalizeCompareValue_(b);
}

// ดึงข้อมูลจากชีตต้นทางมาเทียบกับข้อมูลในระบบ: AssetID ใหม่ที่ยังไม่มี -> เพิ่มเข้ามาและติดป้าย "ข้อมูลใหม่จากบัญชี"
// AssetID ที่มีอยู่แล้วแต่บางฟิลด์ไม่ตรงกับที่มีในระบบ -> "ไม่อัปเดตค่าอัตโนมัติ" แค่ติดป้าย "ไม่ตรงกับบัญชี" พร้อมสรุปรายละเอียดไว้ให้ตรวจสอบเอง
// ป้ายทั้งสองแบบจะหายไปอัตโนมัติเมื่อมีคนแก้ไข/บันทึกทรัพย์สินรายการนั้น (ดู adminSaveAsset_)
function adminSyncFromSource_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const url = getSourceSheetUrl_();
  if (!url) return { ok: false, error: 'ยังไม่ได้ตั้งค่าลิงก์ Google Sheet ต้นทาง' };

  let srcSs;
  try {
    srcSs = SpreadsheetApp.openByUrl(url);
  } catch (err) {
    return { ok: false, error: 'เปิดลิงก์ Google Sheet ไม่สำเร็จ (ตรวจสอบลิงก์และสิทธิ์การเข้าถึง): ' + err };
  }
  const srcSheet = findSourceSheet_(srcSs);
  if (!srcSheet) return { ok: false, error: 'ไม่พบชีตข้อมูลในลิงก์ที่ระบุ' };

  const srcValues = srcSheet.getDataRange().getValues();
  if (srcValues.length < 2) return { ok: false, error: 'ไม่พบข้อมูลในชีตต้นทาง' };
  const srcHeaders = srcValues.shift();
  const srcIdx = indexMap_(srcHeaders);
  if (srcIdx.AssetID === undefined) return { ok: false, error: 'ไม่พบคอลัมน์ AssetID ในชีตต้นทาง' };

  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  const localRowByAssetId = {};
  for (let i = 1; i < values.length; i++) {
    localRowByAssetId[String(values[i][idx.AssetID])] = i + 1;
  }

  let mismatchCount = 0;
  const newRows = [];
  srcValues.forEach(srcRow => {
    const assetId = String(srcRow[srcIdx.AssetID] || '').trim();
    if (!assetId) return;
    const rowNum = localRowByAssetId[assetId];
    if (!rowNum) {
      const newRow = headers.map(() => '');
      newRow[idx.AssetID] = assetId;
      SOURCE_SYNC_FIELDS.forEach(f => {
        if (srcIdx[f] === undefined || idx[f] === undefined) return;
        newRow[idx[f]] = srcRow[srcIdx[f]];
      });
      newRow[idx.UpdatedAt] = new Date();
      newRow[idx.SyncFlag] = 'New';
      newRows.push(newRow);
      return;
    }

    const mismatches = [];
    SOURCE_SYNC_FIELDS.forEach(f => {
      if (srcIdx[f] === undefined || idx[f] === undefined) return;
      const srcVal = srcRow[srcIdx[f]];
      const localVal = values[rowNum - 1][idx[f]];
      if (valuesDiffer_(srcVal, localVal)) {
        mismatches.push((SYNC_FIELD_LABELS_TH[f] || f) + ': บัญชีแจ้ง ' + normalizeCompareValue_(srcVal) + ' / ในระบบ ' + normalizeCompareValue_(localVal));
      }
    });

    if (mismatches.length) {
      mismatchCount++;
      // เขียนป้ายเตือนได้ก็ต่อเมื่อชีตมีคอลัมน์ SyncFlag/SyncNote แล้ว (รัน setup() ใหม่แล้ว)
      if (idx.SyncFlag !== undefined && idx.SyncNote !== undefined) {
        sh.getRange(rowNum, idx.SyncFlag + 1).setValue('Mismatch');
        sh.getRange(rowNum, idx.SyncNote + 1).setValue(mismatches.join('; '));
      }
    } else if (idx.SyncFlag !== undefined && String(values[rowNum - 1][idx.SyncFlag] || '') === 'Mismatch') {
      sh.getRange(rowNum, idx.SyncFlag + 1).setValue('');
      if (idx.SyncNote !== undefined) sh.getRange(rowNum, idx.SyncNote + 1).setValue('');
    }
  });

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  logActivity_('', 'ADMIN_SYNC_SOURCE', 'admin', 'ดึงข้อมูลจากลิงก์ต้นทาง (' + srcSheet.getName() + ') พบข้อมูลใหม่ ' + newRows.length + ' รายการ พบไม่ตรงกับบัญชี ' + mismatchCount + ' รายการ');

  return { ok: true, data: { mismatchCount, addedCount: newRows.length, sourceTabName: srcSheet.getName() } };
}

// ============================================================
// DEPT CODES
// ============================================================
function getDeptCodes_() {
  const sh = getSS_().getSheetByName(SHEETS.DEPT_CODES);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  return values.filter(r => r[idx.DeptName]).map(r => rowToObj_(r, idx));
}

function adminSaveDept_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const dept = body.dept || {};
  const deptName = String(dept.DeptName || '').trim();
  const code = String(dept.Code || '').trim();
  if (!deptName || !code) return { ok: false, error: 'กรุณาระบุชื่อหน่วยงานและรหัส' };
  const sh = getSS_().getSheetByName(SHEETS.DEPT_CODES);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  let rowNum = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.DeptName]) === deptName) { rowNum = i + 1; break; }
  }
  const toSeq = v => (v === undefined || v === null || v === '') ? '' : (Number(v) || '');
  const fields = {
    DeptName: deptName, Code: code,
    ApproverName: String(dept.ApproverName || ''), ApproverEmail: String(dept.ApproverEmail || ''),
    SkipApprovalEmail: !!dept.SkipApprovalEmail,
    StartSeqTransfer: toSeq(dept.StartSeqTransfer), StartSeqSale: toSeq(dept.StartSeqSale), StartSeqWriteOff: toSeq(dept.StartSeqWriteOff)
  };
  if (rowNum === -1) {
    sh.appendRow(HEADERS.DEPT_CODES.map(h => fields[h]));
    logActivity_('', 'ADMIN_SAVE_DEPT', 'admin', 'เพิ่มหน่วยงาน ' + deptName);
    return { ok: true, data: { created: true } };
  }
  Object.keys(fields).forEach(f => sh.getRange(rowNum, idx[f] + 1).setValue(fields[f]));
  logActivity_('', 'ADMIN_SAVE_DEPT', 'admin', 'แก้ไขหน่วยงาน ' + deptName);
  return { ok: true, data: { created: false } };
}

function adminDeleteDept_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const deptName = String(body.deptName || '').trim();
  if (!deptName) return { ok: false, error: 'กรุณาระบุชื่อหน่วยงาน' };
  const sh = getSS_().getSheetByName(SHEETS.DEPT_CODES);
  const values = sh.getDataRange().getValues();
  const idx = indexMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.DeptName]) === deptName) {
      sh.deleteRow(i + 1);
      logActivity_('', 'ADMIN_DELETE_DEPT', 'admin', 'ลบหน่วยงาน ' + deptName);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบหน่วยงานนี้' };
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

  // attach item count + รวมลิงก์รูปภาพของทุกรายการในเอกสารนี้ (สำหรับ export)
  const itemSh = getSS_().getSheetByName(SHEETS.ITEMS);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  const counts = {};
  const images = {};
  itemValues.forEach(r => {
    const tid = r[itemIdx.TransferID];
    counts[tid] = (counts[tid] || 0) + 1;
    const imgs = cellToImages_(r[itemIdx.ImageURL]);
    if (imgs.length) images[tid] = (images[tid] || []).concat(imgs);
  });
  rows.forEach(r => { r.ItemCount = counts[r.TransferID] || 0; r.AllImages = (images[r.TransferID] || []).join(', '); });

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
    .sort((a, b) => a.LineNo - b.LineNo)
    .map(r => { r.Images = cellToImages_(r.ImageURL); return r; });
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

  const dept = getDeptByName_(body.fromDept);
  const skipApproval = !!(dept && dept.SkipApprovalEmail);
  if (!skipApproval && !body.approverEmail) return { ok: false, error: 'ต้องระบุอีเมลผู้อนุมัติ' };

  const fromDeptCode = getCodeForDept_(body.fromDept) || 'GEN';
  const transferId = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();
  const status = skipApproval ? STATUS.APPROVED : STATUS.PENDING;

  // ล็อกช่วงออกเลขที่เอกสาร + บันทึกแถวหลัก กันเลขที่ซ้ำเมื่อมีคนสร้างเอกสารพร้อมกันหลายคน
  const runningNo = withLock_(() => {
    const rn = getNextRunningNo_(fromDeptCode, SHEETS.TRANSFERS, dept && dept.StartSeqTransfer);
    const tSheet = getSS_().getSheetByName(SHEETS.TRANSFERS);
    tSheet.appendRow([
      transferId,
      rn,
      now,
      body.subject || 'โอนย้าย',
      body.subjectOther || '',
      body.purpose || '',
      body.fromDept || '',
      fromDeptCode,
      body.toDept || '',
      status,
      body.approverName || '',
      body.approverEmail || '',
      token,
      skipApproval ? now : '',
      skipApproval ? 'อนุมัติอัตโนมัติ (หน่วยงานนี้ไม่ต้องขออนุมัติ)' : '',
      body.createdBy || '',
      body.createdByEmail || ''
    ]);
    return rn;
  });

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
    imagesToCell_(it.images)
  ]);
  if (itemRows.length) {
    iSheet.getRange(iSheet.getLastRow() + 1, 1, itemRows.length, HEADERS.ITEMS.length).setValues(itemRows);
  }

  let emailResult = { ok: true };
  if (skipApproval) {
    applyTransferToAssets_(transferId);
    logActivity_(transferId, 'CREATE', body.createdBy || 'unknown', 'สร้างใบโอนย้าย ' + runningNo + ' (อนุมัติอัตโนมัติ)');
  } else {
    logActivity_(transferId, 'CREATE', body.createdBy || 'unknown', 'สร้างใบโอนย้าย ' + runningNo);
    emailResult = sendApprovalEmail_(transferId, runningNo, body, items, token);
  }

  return { ok: true, data: { transferId, runningNo, emailSent: emailResult.ok, emailError: emailResult.error || null, autoApproved: skipApproval } };
}

function getCodeForDept_(deptName) {
  const codes = getDeptCodes_();
  const found = codes.find(c => c.DeptName === deptName);
  return found ? found.Code : null;
}

function getDeptByName_(deptName) {
  return getDeptCodes_().find(d => d.DeptName === deptName) || null;
}

// startSeq (ถ้ากำหนดไว้ใน DeptCodes) คือเลขลำดับเริ่มต้นของหน่วยงานนั้นสำหรับเอกสารประเภทนี้
// เลขที่ออกจริงจะเป็นค่ามากกว่าระหว่าง startSeq กับเลขลำดับสูงสุดที่มีอยู่แล้ว +1 เสมอ (ไม่ทับเลขที่ออกไปแล้ว)
function getNextRunningNo_(code, sheetName, startSeq) {
  const buddhistYear = new Date().getFullYear() + 543;
  const yy = String(buddhistYear).slice(-2);
  const sh = getSS_().getSheetByName(sheetName || SHEETS.TRANSFERS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  const prefix = code + '/';
  const suffix = '/' + yy;
  let maxSeq = (startSeq && Number(startSeq) > 0) ? Number(startSeq) - 1 : 0;
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

  if (decision === STATUS.APPROVED) {
    applyTransferToAssets_(transferId);
  }

  logActivity_(transferId, decision.toUpperCase(), found.obj.ApproverName || found.obj.ApproverEmail, body.comment || '');

  sendDecisionNotification_(found.obj, decision, body.comment || '');

  return { ok: true, data: { transferId, status: decision } };
}

// เมื่อใบโอนย้ายได้รับอนุมัติแล้ว ปรับหน่วยงาน/ผู้ดูแลของทรัพย์สินแต่ละรายการในชีต Assets ให้ตรงกับปลายทาง
function applyTransferToAssets_(transferId) {
  const items = getTransferItems_(transferId);
  if (!items.length) return;
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  const rowByAssetId = {};
  for (let i = 1; i < values.length; i++) {
    rowByAssetId[String(values[i][idx.AssetID])] = i + 1;
  }
  items.forEach(it => {
    const rowNum = rowByAssetId[String(it.AssetID)];
    if (!rowNum) return;
    if (it.ToDeptName) sh.getRange(rowNum, idx.Department + 1).setValue(it.ToDeptName);
    if (it.ToSignName) sh.getRange(rowNum, idx.Custodian + 1).setValue(it.ToSignName);
    sh.getRange(rowNum, idx.UpdatedAt + 1).setValue(new Date());
  });
}

// ============================================================
// SALES — ขายออกทรัพย์สิน (ราคาซาก / ราคาประมูล / ราคาขาย)
// ============================================================
function getSales_(params) {
  const sh = getSS_().getSheetByName(SHEETS.SALES);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  let rows = values.map(r => rowToObj_(r, idx));

  if (params.status) rows = rows.filter(r => r.Status === params.status);
  if (params.dept) rows = rows.filter(r => r.FromDept === params.dept);
  if (params.from) rows = rows.filter(r => new Date(r.CreatedAt) >= new Date(params.from));
  if (params.to) rows = rows.filter(r => new Date(r.CreatedAt) <= new Date(params.to + 'T23:59:59'));
  if (params.q) {
    const qq = params.q.toLowerCase();
    rows = rows.filter(r =>
      String(r.RunningNo).toLowerCase().indexOf(qq) !== -1 ||
      String(r.Buyer).toLowerCase().indexOf(qq) !== -1
    );
  }
  rows.sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));

  const itemSh = getSS_().getSheetByName(SHEETS.SALE_ITEMS);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  const counts = {};
  const totals = {};
  const images = {};
  itemValues.forEach(r => {
    const sid = r[itemIdx.SaleID];
    counts[sid] = (counts[sid] || 0) + 1;
    totals[sid] = (totals[sid] || 0) + (parseFloat(r[itemIdx.SalePrice]) || 0);
    const imgs = cellToImages_(r[itemIdx.ImageURL]);
    if (imgs.length) images[sid] = (images[sid] || []).concat(imgs);
  });
  rows.forEach(r => { r.ItemCount = counts[r.SaleID] || 0; r.TotalSalePrice = totals[r.SaleID] || 0; r.AllImages = (images[r.SaleID] || []).join(', '); });

  return rows;
}

function getSaleFull_(saleId) {
  const s = findSaleRow_(saleId);
  if (!s) return null;
  const obj = s.obj;
  obj.Items = getSaleItems_(saleId);
  delete obj.ApprovalToken;
  return obj;
}

function getSaleItems_(saleId) {
  const sh = getSS_().getSheetByName(SHEETS.SALE_ITEMS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  return values
    .map(r => rowToObj_(r, idx))
    .filter(r => r.SaleID === saleId)
    .sort((a, b) => a.LineNo - b.LineNo)
    .map(r => { r.Images = cellToImages_(r.ImageURL); return r; });
}

function findSaleRow_(saleId) {
  const sh = getSS_().getSheetByName(SHEETS.SALES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.SaleID]) === String(saleId)) {
      return { rowNum: i + 1, obj: rowToObj_(values[i], idx), idx };
    }
  }
  return null;
}

function createSale_(body) {
  const items = body.items || [];
  if (!items.length) return { ok: false, error: 'ต้องมีรายการทรัพย์สินอย่างน้อย 1 รายการ' };

  const dept = getDeptByName_(body.fromDept);
  const skipApproval = !!(dept && dept.SkipApprovalEmail);
  if (!skipApproval && !body.approverEmail) return { ok: false, error: 'ต้องระบุอีเมลผู้อนุมัติ' };

  const fromDeptCode = getCodeForDept_(body.fromDept) || 'GEN';
  const saleId = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();
  const status = skipApproval ? STATUS.APPROVED : STATUS.PENDING;

  // ล็อกช่วงออกเลขที่เอกสาร + บันทึกแถวหลัก กันเลขที่ซ้ำเมื่อมีคนสร้างเอกสารพร้อมกันหลายคน
  const runningNo = withLock_(() => {
    const rn = getNextRunningNo_(fromDeptCode + 'S', SHEETS.SALES, dept && dept.StartSeqSale);
    const sSheet = getSS_().getSheetByName(SHEETS.SALES);
    sSheet.appendRow([
      saleId,
      rn,
      now,
      body.fromDept || '',
      fromDeptCode,
      body.buyer || '',
      body.remark || '',
      status,
      body.approverName || '',
      body.approverEmail || '',
      token,
      skipApproval ? now : '',
      skipApproval ? 'อนุมัติอัตโนมัติ (หน่วยงานนี้ไม่ต้องขออนุมัติ)' : '',
      body.createdBy || '',
      body.createdByEmail || ''
    ]);
    return rn;
  });

  const iSheet = getSS_().getSheetByName(SHEETS.SALE_ITEMS);
  const itemRows = items.map((it, i) => [
    saleId,
    i + 1,
    it.assetId || '',
    it.assetName || '',
    it.scrapPrice || 0,
    it.auctionPrice || 0,
    it.salePrice || 0,
    it.remark || '',
    imagesToCell_(it.images)
  ]);
  if (itemRows.length) {
    iSheet.getRange(iSheet.getLastRow() + 1, 1, itemRows.length, HEADERS.SALE_ITEMS.length).setValues(itemRows);
  }

  let emailResult = { ok: true };
  if (skipApproval) {
    const soldAssetIdsNow = items.map(it => it.assetId).filter(Boolean);
    purgeAssetFromAllQueues_(soldAssetIdsNow);
    setAssetsTag_(soldAssetIdsNow, 'ขาย');
    logActivity_(saleId, 'CREATE_SALE', body.createdBy || 'unknown', 'สร้างใบขายออกทรัพย์สิน ' + runningNo + ' (อนุมัติอัตโนมัติ)');
  } else {
    logActivity_(saleId, 'CREATE_SALE', body.createdBy || 'unknown', 'สร้างใบขายออกทรัพย์สิน ' + runningNo);
    emailResult = sendSaleApprovalEmail_(saleId, runningNo, body, items, token);
  }

  return { ok: true, data: { saleId, runningNo, emailSent: emailResult.ok, emailError: emailResult.error || null, autoApproved: skipApproval } };
}

function getSaleApprovalView_(saleId, token) {
  const found = findSaleRow_(saleId);
  if (!found) return { ok: false, error: 'ไม่พบใบขายออกนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  const items = getSaleItems_(saleId);
  const obj = Object.assign({}, found.obj);
  delete obj.ApprovalToken;
  obj.Items = items;
  return { ok: true, data: obj };
}

function decideSale_(body) {
  const saleId = body.saleId;
  const token = body.token;
  const decision = body.decision; // 'Approved' | 'Rejected'
  const found = findSaleRow_(saleId);
  if (!found) return { ok: false, error: 'ไม่พบใบขายออกนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  if (found.obj.Status !== STATUS.PENDING) {
    return { ok: false, error: 'ใบขายออกนี้ถูกดำเนินการไปแล้ว (สถานะปัจจุบัน: ' + found.obj.Status + ')' };
  }
  if (decision !== STATUS.APPROVED && decision !== STATUS.REJECTED) {
    return { ok: false, error: 'decision ไม่ถูกต้อง' };
  }

  const sh = getSS_().getSheetByName(SHEETS.SALES);
  const idx = found.idx;
  const rowNum = found.rowNum;
  sh.getRange(rowNum, idx.Status + 1).setValue(decision);
  sh.getRange(rowNum, idx.ApprovedAt + 1).setValue(new Date());
  sh.getRange(rowNum, idx.ApproverComment + 1).setValue(body.comment || '');

  if (decision === STATUS.APPROVED) {
    const soldAssetIds = getSaleItems_(saleId).map(it => it.AssetID).filter(Boolean);
    purgeAssetFromAllQueues_(soldAssetIds);
    setAssetsTag_(soldAssetIds, 'ขาย');
  }

  logActivity_(saleId, 'SALE_' + decision.toUpperCase(), found.obj.ApproverName || found.obj.ApproverEmail, body.comment || '');

  sendSaleDecisionNotification_(found.obj, decision, body.comment || '');

  return { ok: true, data: { saleId, status: decision } };
}

// ============================================================
// WRITE-OFFS — ตัดชำรุดทรัพย์สิน (ราคาซาก)
// ============================================================
function getWriteOffs_(params) {
  const sh = getSS_().getSheetByName(SHEETS.WRITEOFFS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  let rows = values.map(r => rowToObj_(r, idx));

  if (params.status) rows = rows.filter(r => r.Status === params.status);
  if (params.dept) rows = rows.filter(r => r.FromDept === params.dept);
  if (params.from) rows = rows.filter(r => new Date(r.CreatedAt) >= new Date(params.from));
  if (params.to) rows = rows.filter(r => new Date(r.CreatedAt) <= new Date(params.to + 'T23:59:59'));
  if (params.q) {
    const qq = params.q.toLowerCase();
    rows = rows.filter(r =>
      String(r.RunningNo).toLowerCase().indexOf(qq) !== -1 ||
      String(r.Reason).toLowerCase().indexOf(qq) !== -1
    );
  }
  rows.sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));

  const itemSh = getSS_().getSheetByName(SHEETS.WRITEOFF_ITEMS);
  const itemValues = itemSh.getDataRange().getValues();
  const itemHeaders = itemValues.shift();
  const itemIdx = indexMap_(itemHeaders);
  const counts = {};
  const images = {};
  itemValues.forEach(r => {
    const wid = r[itemIdx.WriteOffID];
    counts[wid] = (counts[wid] || 0) + 1;
    const imgs = cellToImages_(r[itemIdx.ImageURL]);
    if (imgs.length) images[wid] = (images[wid] || []).concat(imgs);
  });
  rows.forEach(r => { r.ItemCount = counts[r.WriteOffID] || 0; r.AllImages = (images[r.WriteOffID] || []).join(', '); });

  return rows;
}

function getWriteOffFull_(writeOffId) {
  const w = findWriteOffRow_(writeOffId);
  if (!w) return null;
  const obj = w.obj;
  obj.Items = getWriteOffItems_(writeOffId);
  delete obj.ApprovalToken;
  return obj;
}

function getWriteOffItems_(writeOffId) {
  const sh = getSS_().getSheetByName(SHEETS.WRITEOFF_ITEMS);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = indexMap_(headers);
  return values
    .map(r => rowToObj_(r, idx))
    .filter(r => r.WriteOffID === writeOffId)
    .sort((a, b) => a.LineNo - b.LineNo)
    .map(r => { r.Images = cellToImages_(r.ImageURL); return r; });
}

function findWriteOffRow_(writeOffId) {
  const sh = getSS_().getSheetByName(SHEETS.WRITEOFFS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.WriteOffID]) === String(writeOffId)) {
      return { rowNum: i + 1, obj: rowToObj_(values[i], idx), idx };
    }
  }
  return null;
}

function createWriteOff_(body) {
  const items = body.items || [];
  if (!items.length) return { ok: false, error: 'ต้องมีรายการทรัพย์สินอย่างน้อย 1 รายการ' };

  const dept = getDeptByName_(body.fromDept);
  const skipApproval = !!(dept && dept.SkipApprovalEmail);
  if (!skipApproval && !body.approverEmail) return { ok: false, error: 'ต้องระบุอีเมลผู้อนุมัติ' };

  const fromDeptCode = getCodeForDept_(body.fromDept) || 'GEN';
  const writeOffId = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();
  const status = skipApproval ? STATUS.APPROVED : STATUS.PENDING;

  // ล็อกช่วงออกเลขที่เอกสาร + บันทึกแถวหลัก กันเลขที่ซ้ำเมื่อมีคนสร้างเอกสารพร้อมกันหลายคน
  const runningNo = withLock_(() => {
    const rn = getNextRunningNo_(fromDeptCode + 'W', SHEETS.WRITEOFFS, dept && dept.StartSeqWriteOff);
    const wSheet = getSS_().getSheetByName(SHEETS.WRITEOFFS);
    wSheet.appendRow([
      writeOffId,
      rn,
      now,
      body.fromDept || '',
      fromDeptCode,
      body.reason || '',
      body.remark || '',
      status,
      body.approverName || '',
      body.approverEmail || '',
      token,
      skipApproval ? now : '',
      skipApproval ? 'อนุมัติอัตโนมัติ (หน่วยงานนี้ไม่ต้องขออนุมัติ)' : '',
      body.createdBy || '',
      body.createdByEmail || ''
    ]);
    return rn;
  });

  const iSheet = getSS_().getSheetByName(SHEETS.WRITEOFF_ITEMS);
  const itemRows = items.map((it, i) => [
    writeOffId,
    i + 1,
    it.assetId || '',
    it.assetName || '',
    it.scrapPrice || 0,
    it.remark || '',
    imagesToCell_(it.images)
  ]);
  if (itemRows.length) {
    iSheet.getRange(iSheet.getLastRow() + 1, 1, itemRows.length, HEADERS.WRITEOFF_ITEMS.length).setValues(itemRows);
  }

  let emailResult = { ok: true };
  if (skipApproval) {
    const writtenOffAssetIdsNow = items.map(it => it.assetId).filter(Boolean);
    purgeAssetFromAllQueues_(writtenOffAssetIdsNow);
    setAssetsTag_(writtenOffAssetIdsNow, 'ชำรุด');
    logActivity_(writeOffId, 'CREATE_WRITEOFF', body.createdBy || 'unknown', 'สร้างใบตัดชำรุดทรัพย์สิน ' + runningNo + ' (อนุมัติอัตโนมัติ)');
  } else {
    logActivity_(writeOffId, 'CREATE_WRITEOFF', body.createdBy || 'unknown', 'สร้างใบตัดชำรุดทรัพย์สิน ' + runningNo);
    emailResult = sendWriteOffApprovalEmail_(writeOffId, runningNo, body, items, token);
  }

  return { ok: true, data: { writeOffId, runningNo, emailSent: emailResult.ok, emailError: emailResult.error || null, autoApproved: skipApproval } };
}

function getWriteOffApprovalView_(writeOffId, token) {
  const found = findWriteOffRow_(writeOffId);
  if (!found) return { ok: false, error: 'ไม่พบใบตัดชำรุดนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  const items = getWriteOffItems_(writeOffId);
  const obj = Object.assign({}, found.obj);
  delete obj.ApprovalToken;
  obj.Items = items;
  return { ok: true, data: obj };
}

function decideWriteOff_(body) {
  const writeOffId = body.writeOffId;
  const token = body.token;
  const decision = body.decision; // 'Approved' | 'Rejected'
  const found = findWriteOffRow_(writeOffId);
  if (!found) return { ok: false, error: 'ไม่พบใบตัดชำรุดนี้' };
  if (String(found.obj.ApprovalToken) !== String(token)) {
    return { ok: false, error: 'ลิงก์ไม่ถูกต้องหรือหมดอายุ' };
  }
  if (found.obj.Status !== STATUS.PENDING) {
    return { ok: false, error: 'ใบตัดชำรุดนี้ถูกดำเนินการไปแล้ว (สถานะปัจจุบัน: ' + found.obj.Status + ')' };
  }
  if (decision !== STATUS.APPROVED && decision !== STATUS.REJECTED) {
    return { ok: false, error: 'decision ไม่ถูกต้อง' };
  }

  const sh = getSS_().getSheetByName(SHEETS.WRITEOFFS);
  const idx = found.idx;
  const rowNum = found.rowNum;
  sh.getRange(rowNum, idx.Status + 1).setValue(decision);
  sh.getRange(rowNum, idx.ApprovedAt + 1).setValue(new Date());
  sh.getRange(rowNum, idx.ApproverComment + 1).setValue(body.comment || '');

  if (decision === STATUS.APPROVED) {
    const writtenOffAssetIds = getWriteOffItems_(writeOffId).map(it => it.AssetID).filter(Boolean);
    purgeAssetFromAllQueues_(writtenOffAssetIds);
    setAssetsTag_(writtenOffAssetIds, 'ชำรุด');
  }

  logActivity_(writeOffId, 'WRITEOFF_' + decision.toUpperCase(), found.obj.ApproverName || found.obj.ApproverEmail, body.comment || '');

  sendWriteOffDecisionNotification_(found.obj, decision, body.comment || '');

  return { ok: true, data: { writeOffId, status: decision } };
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
      '<td style="border:1px solid #ddd;padding:6px;">' + ((it.images || []).filter(Boolean).slice(0, ITEM_IMAGE_LIMIT).map(u => '<img src="' + u + '" width="34" style="border-radius:4px;margin:1px;">').join('') || '-') + '</td>' +
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

function sendSaleApprovalEmail_(saleId, runningNo, body, items, token) {
  try {
    const approveUrl = CONFIG.FRONTEND_URL + '?view=approveSale&id=' + encodeURIComponent(saleId) + '&token=' + encodeURIComponent(token);
    const itemsHtml = items.map((it, i) => (
      '<tr>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">' + (i + 1) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetId) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetName) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:right;">' + fmtMoneyServer_(it.scrapPrice) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:right;">' + fmtMoneyServer_(it.auctionPrice) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:right;">' + fmtMoneyServer_(it.salePrice) + '</td>' +
      '</tr>'
    )).join('');

    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:640px;margin:auto;">' +
      '<h2 style="color:#1a3c6e;">' + CONFIG.COMPANY_NAME + '</h2>' +
      '<h3>ใบขายออกทรัพย์สิน เลขที่ ' + runningNo + '</h3>' +
      '<p><b>หน่วยงาน:</b> ' + escapeHtml_(body.fromDept) + '</p>' +
      '<p><b>ผู้ซื้อ/ผู้ประมูลได้:</b> ' + escapeHtml_(body.buyer || '-') + '</p>' +
      '<p><b>หมายเหตุ:</b> ' + escapeHtml_(body.remark || '-') + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
      '<tr style="background:#f0f4f8;"><th style="border:1px solid #ddd;padding:6px;">#</th><th style="border:1px solid #ddd;padding:6px;">รหัส</th><th style="border:1px solid #ddd;padding:6px;">รายการ</th><th style="border:1px solid #ddd;padding:6px;">ราคาซาก</th><th style="border:1px solid #ddd;padding:6px;">ราคาประมูล</th><th style="border:1px solid #ddd;padding:6px;">ราคาขาย</th></tr>' +
      itemsHtml +
      '</table>' +
      '<p style="margin-top:20px;">กรุณาตรวจสอบรายละเอียดฉบับเต็มและพิจารณาอนุมัติที่ลิงก์ด้านล่าง:</p>' +
      '<p><a href="' + approveUrl + '" style="background:#1a3c6e;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;display:inline-block;">เปิดใบขายออกเพื่อพิจารณาอนุมัติ</a></p>' +
      '<p style="color:#888;font-size:12px;">ผู้บันทึก: ' + escapeHtml_(body.createdBy || '-') + '</p>' +
      '</div>';

    MailApp.sendEmail({
      to: body.approverEmail,
      subject: '[ขออนุมัติ] ใบขายออกทรัพย์สิน ' + runningNo + ' — ' + CONFIG.COMPANY_NAME,
      htmlBody: html
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function sendSaleDecisionNotification_(saleObj, decision, comment) {
  try {
    if (!saleObj.CreatedByEmail) return;
    const statusThai = decision === STATUS.APPROVED ? 'ขายแล้ว' : 'ไม่อนุมัติ';
    const color = decision === STATUS.APPROVED ? '#1a7d3c' : '#c0392b';
    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:auto;">' +
      '<h3>ใบขายออกทรัพย์สิน เลขที่ ' + saleObj.RunningNo + '</h3>' +
      '<p style="font-size:16px;">สถานะ: <b style="color:' + color + ';">' + statusThai + '</b></p>' +
      (comment ? '<p><b>ความเห็นผู้อนุมัติ:</b> ' + escapeHtml_(comment) + '</p>' : '') +
      '<p>โดย: ' + escapeHtml_(saleObj.ApproverName || saleObj.ApproverEmail) + '</p>' +
      '</div>';
    MailApp.sendEmail({
      to: saleObj.CreatedByEmail,
      subject: '[' + statusThai + '] ใบขายออกทรัพย์สิน ' + saleObj.RunningNo,
      htmlBody: html
    });
  } catch (err) {
    Logger.log('sendSaleDecisionNotification_ error: ' + err);
  }
}

function sendWriteOffApprovalEmail_(writeOffId, runningNo, body, items, token) {
  try {
    const approveUrl = CONFIG.FRONTEND_URL + '?view=approveWriteOff&id=' + encodeURIComponent(writeOffId) + '&token=' + encodeURIComponent(token);
    const itemsHtml = items.map((it, i) => (
      '<tr>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">' + (i + 1) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetId) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.assetName) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:right;">' + fmtMoneyServer_(it.scrapPrice) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.remark || '') + '</td>' +
      '</tr>'
    )).join('');

    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:640px;margin:auto;">' +
      '<h2 style="color:#1a3c6e;">' + CONFIG.COMPANY_NAME + '</h2>' +
      '<h3>ใบตัดชำรุดทรัพย์สิน เลขที่ ' + runningNo + '</h3>' +
      '<p><b>หน่วยงาน:</b> ' + escapeHtml_(body.fromDept) + '</p>' +
      '<p><b>สาเหตุชำรุด:</b> ' + escapeHtml_(body.reason || '-') + '</p>' +
      '<p><b>หมายเหตุ:</b> ' + escapeHtml_(body.remark || '-') + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
      '<tr style="background:#f0f4f8;"><th style="border:1px solid #ddd;padding:6px;">#</th><th style="border:1px solid #ddd;padding:6px;">รหัส</th><th style="border:1px solid #ddd;padding:6px;">รายการ</th><th style="border:1px solid #ddd;padding:6px;">ราคาซาก</th><th style="border:1px solid #ddd;padding:6px;">หมายเหตุ</th></tr>' +
      itemsHtml +
      '</table>' +
      '<p style="margin-top:20px;">กรุณาตรวจสอบรายละเอียดฉบับเต็มและพิจารณาอนุมัติที่ลิงก์ด้านล่าง:</p>' +
      '<p><a href="' + approveUrl + '" style="background:#1a3c6e;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;display:inline-block;">เปิดใบตัดชำรุดเพื่อพิจารณาอนุมัติ</a></p>' +
      '<p style="color:#888;font-size:12px;">ผู้บันทึก: ' + escapeHtml_(body.createdBy || '-') + '</p>' +
      '</div>';

    MailApp.sendEmail({
      to: body.approverEmail,
      subject: '[ขออนุมัติ] ใบตัดชำรุดทรัพย์สิน ' + runningNo + ' — ' + CONFIG.COMPANY_NAME,
      htmlBody: html
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function sendWriteOffDecisionNotification_(writeOffObj, decision, comment) {
  try {
    if (!writeOffObj.CreatedByEmail) return;
    const statusThai = decision === STATUS.APPROVED ? 'อนุมัติ' : 'ไม่อนุมัติ';
    const color = decision === STATUS.APPROVED ? '#1a7d3c' : '#c0392b';
    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:auto;">' +
      '<h3>ใบตัดชำรุดทรัพย์สิน เลขที่ ' + writeOffObj.RunningNo + '</h3>' +
      '<p style="font-size:16px;">สถานะ: <b style="color:' + color + ';">' + statusThai + '</b></p>' +
      (comment ? '<p><b>ความเห็นผู้อนุมัติ:</b> ' + escapeHtml_(comment) + '</p>' : '') +
      '<p>โดย: ' + escapeHtml_(writeOffObj.ApproverName || writeOffObj.ApproverEmail) + '</p>' +
      '</div>';
    MailApp.sendEmail({
      to: writeOffObj.CreatedByEmail,
      subject: '[' + statusThai + '] ใบตัดชำรุดทรัพย์สิน ' + writeOffObj.RunningNo,
      htmlBody: html
    });
  } catch (err) {
    Logger.log('sendWriteOffDecisionNotification_ error: ' + err);
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
// กันชนกันตอนมีคนใช้งานพร้อมกันหลายคน (เช่น ออกเลขที่เอกสารซ้ำ) — ล็อกให้ทำงานทีละคนเฉพาะช่วงวิกฤต
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

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

// รายการทรัพย์สินแต่ละรายการแนบรูปได้สูงสุด 5 รูป เก็บใน 1 คอลัมน์ ImageURL คั่นด้วยจุลภาค
const ITEM_IMAGE_LIMIT = 5;
function imagesToCell_(images) {
  return (images || []).filter(Boolean).slice(0, ITEM_IMAGE_LIMIT).join(',');
}
function cellToImages_(cell) {
  return String(cell || '').split(',').map(s => s.trim()).filter(Boolean);
}

function escapeHtml_(str) {
  return String(str || '').replace(/[&<>"']/g, function (m) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
  });
}

function fmtMoneyServer_(n) {
  n = parseFloat(n) || 0;
  return n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
