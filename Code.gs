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
  DEFAULT_TIMEOUT_MS: 15000,
  // SharePoint — ไม่ใช่ความลับ (แค่ระบุตำแหน่งไซต์/โฟลเดอร์) ส่วน Tenant/Client ID/Secret เก็บแยกใน
  // Script Properties เท่านั้น (ดู getSharePointAccessToken_) ไม่ฝังในโค้ดนี้เพื่อความปลอดภัย
  SHAREPOINT_HOSTNAME: 'siammotor.sharepoint.com',
  SHAREPOINT_SITE_PATH: '/sites/Chosiya_Server',
  SHAREPOINT_FOLDER: 'Doc_Assets-2026'
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
  USERS: ['Username', 'Password', 'Role', 'Departments', 'CanViewPrices', 'CreatedAt'],
  TRANSFER_QUEUE: ['AssetID', 'Purpose', 'AddedBy', 'AddedAt'],
  TRANSFERS: ['TransferID', 'RunningNo', 'CreatedAt', 'Subject', 'SubjectOther', 'Purpose', 'FromDept', 'FromDeptCode', 'ToDept', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail', 'NotifiedAt'],
  ITEMS: ['TransferID', 'LineNo', 'AssetID', 'AssetName', 'FromDeptName', 'FromSignName', 'ToDeptName', 'ToSignName', 'Remark', 'ImageURL'],
  SALES: ['SaleID', 'RunningNo', 'CreatedAt', 'FromDept', 'FromDeptCode', 'Buyer', 'Remark', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail', 'NotifiedAt'],
  SALE_ITEMS: ['SaleID', 'LineNo', 'AssetID', 'AssetName', 'ScrapPrice', 'AuctionPrice', 'SalePrice', 'Remark', 'ImageURL'],
  WRITEOFFS: ['WriteOffID', 'RunningNo', 'CreatedAt', 'FromDept', 'FromDeptCode', 'Reason', 'Remark', 'Status', 'ApproverName', 'ApproverEmail', 'ApprovalToken', 'ApprovedAt', 'ApproverComment', 'CreatedBy', 'CreatedByEmail', 'NotifiedAt'],
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
      case 'getAssetListBundle':
        result = { ok: true, data: getAssetListBundle_(e.parameter.q || '') };
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
      case 'adminSyncScrapPriceToBookValue':
        result = adminSyncScrapPriceToBookValue_(body);
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
      case 'adminClearAssetQueues':
        result = adminClearAssetQueues_(body);
        break;
      case 'adminSaveNotifyEmails':
        result = adminSaveNotifyEmails_(body);
        break;
      case 'notifyAccountingGA':
        result = notifyAccountingGA_(body);
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
  // ราคาซาก ใช้ค่าที่บันทึกไว้จริงในคอลัมน์ ScrapPrice ของชีต Assets โดยตรง (rowToObj_ อ่านมาให้แล้วด้านบน)
  // ไม่คำนวณอัตโนมัติจาก % อีกต่อไป — Admin แก้ไขค่านี้ได้ที่หน้า "รายการทรัพย์สิน" > แก้ไขทรัพย์สิน
  return rows;
}

// ============================================================
// ASSET QUEUE — ทรัพย์สินที่มาร์คไว้ล่วงหน้าจากหน้า "รายการทรัพย์สิน" เพื่อรอออกใบโอนย้าย/ขายออก/ตัดชำรุด
// ใช้ชีตเดียวกันร่วมกันทั้ง 3 ประเภท แยกด้วยคอลัมน์ Purpose
// ============================================================
const QUEUE_PURPOSES = { TRANSFER: 'Transfer', SALE: 'Sale', WRITEOFF: 'WriteOff' };
// ทรัพย์สิน 1 รายการ อยู่ในคิวรอได้ทีละ 1 ประเภทเท่านั้น (โอนย้าย/ขาย/ตัดชำรุด เลือกได้อย่างใดอย่างหนึ่ง)
const QUEUE_PURPOSE_LABEL_TH = { Transfer: 'โอนย้าย', Sale: 'ขาย', WriteOff: 'ตัดชำรุด' };

// ประกอบแถวคิวรอ 1 แถวจากข้อมูลคิว (q) + ข้อมูลทรัพย์สินที่เกี่ยวข้อง (a) + สถานะการขาย/ตัดชำรุด (disposed)
// แยกออกมาเป็นฟังก์ชันกลาง เพื่อให้ getAssetQueue_ (ใช้ทีละประเภท) และ getAssetListBundle_ (ใช้รวมทั้ง 3 ประเภทในคราวเดียว)
// ได้ผลลัพธ์รูปแบบเดียวกันเป๊ะ ไม่ต้องคัดลอกโค้ดซ้ำ
function buildQueueRow_(q, a, disposed) {
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
}

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

  return queueRows
    .map(q => buildQueueRow_(q, assetsById[String(q.AssetID)] || {}, disposed))
    .sort((x, y) => new Date(x.AddedAt) - new Date(y.AddedAt));
}

// เวอร์ชันรวมของ getAssetsFull_ + getTransferQueue_/getSaleQueue_/getWriteOffQueue_ สำหรับหน้า "รายการทรัพย์สิน" โดยเฉพาะ
// คำนวณ getDisposedAssetStatus_()/getAssetsRaw_()/อ่านชีต TransferQueue "ครั้งเดียว" แล้วประกอบผลลัพธ์ทั้ง 4 ส่วนจากข้อมูลชุดเดียวกัน
// แทนที่จะให้ 4 request แยกกัน (เดิม) ต่างคนต่างอ่านชีตเดิมซ้ำ — action เดิมทั้ง 4 ยังคงอยู่ตามปกติสำหรับหน้าอื่นที่ใช้แยกกัน
function getAssetListBundle_(q) {
  const disposed = getDisposedAssetStatus_();
  const allAssets = getAssetsRaw_();
  allAssets.forEach(r => { r.AssetStatus = disposed[String(r.AssetID)] || 'Active'; });

  let assets = allAssets;
  if (q) {
    const qq = q.toString().toLowerCase();
    assets = allAssets.filter(r =>
      String(r.AssetID).toLowerCase().indexOf(qq) !== -1 ||
      String(r.AssetName).toLowerCase().indexOf(qq) !== -1 ||
      String(r.Custodian).toLowerCase().indexOf(qq) !== -1
    );
  }

  const assetsById = {};
  allAssets.forEach(a => { assetsById[String(a.AssetID)] = a; });

  const qsh = getSS_().getSheetByName(SHEETS.TRANSFER_QUEUE);
  const qvalues = qsh.getDataRange().getValues();
  const qheaders = qvalues.shift();
  const qidx = indexMap_(qheaders);
  const allQueueRows = qvalues.filter(r => r[qidx.AssetID]).map(r => rowToObj_(r, qidx));

  function queueByPurpose(purpose) {
    return allQueueRows
      .filter(r => (r.Purpose || QUEUE_PURPOSES.TRANSFER) === purpose)
      .map(r => buildQueueRow_(r, assetsById[String(r.AssetID)] || {}, disposed))
      .sort((x, y) => new Date(x.AddedAt) - new Date(y.AddedAt));
  }

  return {
    assets: assets,
    transferQueue: queueByPurpose(QUEUE_PURPOSES.TRANSFER),
    saleQueue: queueByPurpose(QUEUE_PURPOSES.SALE),
    writeOffQueue: queueByPurpose(QUEUE_PURPOSES.WRITEOFF)
  };
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
      if (String(values[i][idx.AssetID]) !== assetId) continue;
      const rowPurpose = idx.Purpose !== undefined ? (values[i][idx.Purpose] || QUEUE_PURPOSES.TRANSFER) : QUEUE_PURPOSES.TRANSFER;
      if (rowPurpose === purpose) return { ok: true, data: { alreadyQueued: true } };
      // ทรัพย์สินนี้อยู่ในคิวรอประเภทอื่นอยู่แล้ว — เลือกได้เพียงคิวเดียวต่อทรัพย์สิน
      return { ok: false, error: 'ทรัพย์สินนี้อยู่ในคิวรอ' + (QUEUE_PURPOSE_LABEL_TH[rowPurpose] || rowPurpose) + 'อยู่แล้ว เลือกได้เพียงคิวเดียวต่อทรัพย์สิน' };
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

// Admin ใช้ล้างคิวรอที่ค้างซ้ำมากกว่า 1 ประเภทของทรัพย์สินชิ้นเดียว (ข้อมูลเก่าก่อนมีการบังคับ "เลือกได้คิวเดียว"
// ใน addToAssetQueue_ หรือกรณีผิดปกติอื่น) — เอาออกจากคิวรอทุกประเภท ให้เลือกเข้าคิวใหม่ได้ถูกต้องอีกครั้ง
function adminClearAssetQueues_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const assetId = String(body.assetId || '').trim();
  if (!assetId) return { ok: false, error: 'กรุณาระบุรหัสทรัพย์สิน' };
  purgeAssetFromAllQueues_([assetId]);
  logActivity_('', 'ADMIN_CLEAR_ASSET_QUEUES', 'admin', 'ล้างคิวรอซ้ำของทรัพย์สิน ' + assetId);
  return { ok: true };
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
        departments: parseDepartments_(values[i][idx.Departments]),
        canViewPrices: values[i][idx.Role] === 'admin' || String(values[i][idx.CanViewPrices]).toLowerCase() === 'true'
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
        departments: parseDepartments_(values[i][idx.Departments]),
        canViewPrices: String(values[i][idx.Role]) === 'admin' || String(values[i][idx.CanViewPrices]).toLowerCase() === 'true'
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
    CanViewPrices: String(r[idx.CanViewPrices]).toLowerCase() === 'true',
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
  const canViewPrices = u.CanViewPrices ? 'true' : 'false';

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
    // ถ้าชีตยังไม่มีคอลัมน์ Departments/CanViewPrices (ยังไม่ได้รัน setup() ใหม่)
    const newRow = headers.map(() => '');
    newRow[idx.Username] = username;
    newRow[idx.Password] = newPassword;
    newRow[idx.Role] = role;
    if (idx.Departments !== undefined) newRow[idx.Departments] = departments;
    if (idx.CanViewPrices !== undefined) newRow[idx.CanViewPrices] = canViewPrices;
    newRow[idx.CreatedAt] = new Date();
    sh.appendRow(newRow);
    logActivity_('', 'ADMIN_SAVE_USER', 'admin', 'เพิ่มผู้ใช้ ' + username);
    return { ok: true, data: { created: true } };
  }
  sh.getRange(rowNum, idx.Role + 1).setValue(role);
  if (idx.Departments !== undefined) sh.getRange(rowNum, idx.Departments + 1).setValue(departments);
  if (idx.CanViewPrices !== undefined) sh.getRange(rowNum, idx.CanViewPrices + 1).setValue(canViewPrices);
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

// ScrapPrice ไม่อยู่ในนี้แล้ว เพราะบังคับให้เท่ากับ BookValue เสมอ (ดูท้าย adminSaveAsset_)
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
      if (h === 'ScrapPrice') return asset.BookValue || '';
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
  // ราคาซาก = มูลค่าตามบัญชีเสมอ (ไม่รับค่าที่พิมพ์แยก) — ใช้ค่าที่เพิ่งบันทึกถ้ามี ไม่งั้นใช้ค่าเดิมในชีต
  if (idx.ScrapPrice !== undefined) {
    const bookValue = asset.BookValue !== undefined ? asset.BookValue : values[rowNum - 1][idx.BookValue];
    sh.getRange(rowNum, idx.ScrapPrice + 1).setValue(bookValue);
  }
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

// ราคาซาก (ScrapPrice) ถูกบังคับให้เท่ากับมูลค่าตามบัญชี (BookValue) เสมอ (ดู adminSaveAsset_/adminSyncFromSource_)
// ใช้ adminSyncScrapPriceToBookValue_ เพื่อไล่แก้ข้อมูลเก่าที่ยังไม่ตรงกันย้อนหลัง
function adminSyncScrapPriceToBookValue_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const sh = getSS_().getSheetByName(SHEETS.ASSETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);
  if (idx.ScrapPrice === undefined || idx.BookValue === undefined) {
    return { ok: false, error: 'ไม่พบคอลัมน์ ScrapPrice หรือ BookValue ในชีต Assets กรุณารัน setup() ใหม่' };
  }
  let fixedCount = 0;
  for (let i = 1; i < values.length; i++) {
    const bookValue = values[i][idx.BookValue] || 0;
    const scrapPrice = values[i][idx.ScrapPrice] || 0;
    if (String(bookValue) !== String(scrapPrice) && parseFloat(bookValue) !== parseFloat(scrapPrice)) {
      sh.getRange(i + 1, idx.ScrapPrice + 1).setValue(bookValue);
      fixedCount++;
    }
  }
  logActivity_('', 'ADMIN_SYNC_SCRAPPRICE', 'admin', 'ซิงค์ราคาซาก = มูลค่าตามบัญชีย้อนหลัง ' + fixedCount + ' รายการ');
  return { ok: true, data: { fixedCount } };
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
  const notifyEmails = getNotifyEmails_();
  return { ok: true, data: { sourceSheetUrl: getSourceSheetUrl_(), accountingEmail: notifyEmails.accounting, gaEmail: notifyEmails.ga } };
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

// ============================================================
// NOTIFY ACCOUNTING/GA — ปุ่ม "แจ้งบัญชี/GA" ในตารางรายการใบโอนย้าย/ขายออก/ตัดชำรุด แสดงเฉพาะเอกสารที่อนุมัติแล้ว
// อีเมลปลายทางตั้งค่าได้จากหน้า "ตั้งค่า" (ไม่ผูกกับ ApproverEmail ของหน่วยงานใดหน่วยงานหนึ่ง เพราะฝ่ายบัญชี/GA
// ต้องได้รับแจ้งทุกเอกสารโดยไม่ขึ้นกับว่าใครเป็นผู้อนุมัติ)
const ACCOUNTING_NOTIFY_EMAIL_PROP = 'ACCOUNTING_NOTIFY_EMAIL';
const GA_NOTIFY_EMAIL_PROP = 'GA_NOTIFY_EMAIL';

function getNotifyEmails_() {
  const props = PropertiesService.getScriptProperties();
  return {
    accounting: props.getProperty(ACCOUNTING_NOTIFY_EMAIL_PROP) || '',
    ga: props.getProperty(GA_NOTIFY_EMAIL_PROP) || ''
  };
}

function adminSaveNotifyEmails_(body) {
  if (!checkAdminPassword_(body.password)) return { ok: false, error: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  const accounting = String(body.accountingEmail || '').trim();
  const ga = String(body.gaEmail || '').trim();
  PropertiesService.getScriptProperties().setProperty(ACCOUNTING_NOTIFY_EMAIL_PROP, accounting);
  PropertiesService.getScriptProperties().setProperty(GA_NOTIFY_EMAIL_PROP, ga);
  logActivity_('', 'ADMIN_SET_NOTIFY_EMAILS', 'admin', 'ตั้งค่าอีเมลแจ้งเตือนฝ่ายบัญชี/GA');
  return { ok: true };
}

const NOTIFY_DOC_LABELS_TH = { transfer: 'ใบโอนย้ายทรัพย์สิน', sale: 'ใบขายออกทรัพย์สิน', writeoff: 'ใบตัดชำรุดทรัพย์สิน' };

function notifyAccountingGA_(body) {
  const docType = body.docType;
  const id = String(body.id || '').trim();
  const docLabel = NOTIFY_DOC_LABELS_TH[docType];
  if (!docLabel || !id) return { ok: false, error: 'ข้อมูลเอกสารไม่ถูกต้อง' };

  const emails = getNotifyEmails_();
  const recipients = [emails.accounting, emails.ga].filter(Boolean);
  if (!recipients.length) return { ok: false, error: 'ยังไม่ได้ตั้งค่าอีเมลฝ่ายบัญชี/GA กรุณาตั้งค่าในหน้า "ตั้งค่า" ก่อน' };

  let found, items, sheetName;
  if (docType === 'transfer') { found = findTransferRow_(id); items = found ? getTransferItems_(id) : []; sheetName = SHEETS.TRANSFERS; }
  else if (docType === 'sale') { found = findSaleRow_(id); items = found ? getSaleItems_(id) : []; sheetName = SHEETS.SALES; }
  else { found = findWriteOffRow_(id); items = found ? getWriteOffItems_(id) : []; sheetName = SHEETS.WRITEOFFS; }
  if (!found) return { ok: false, error: 'ไม่พบเอกสารนี้' };
  if (found.obj.Status !== STATUS.APPROVED) return { ok: false, error: 'แจ้งเตือนได้เฉพาะเอกสารที่อนุมัติแล้วเท่านั้น' };

  try {
    const itemsHtml = items.map((it, i) => (
      '<tr>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">' + (i + 1) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.AssetID) + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;">' + escapeHtml_(it.AssetName) + '</td>' +
      '</tr>'
    )).join('');
    const html =
      '<div style="font-family:Sarabun,Arial,sans-serif;max-width:640px;margin:auto;">' +
      '<h2 style="color:#1a3c6e;">' + escapeHtml_(CONFIG.COMPANY_NAME) + '</h2>' +
      '<h3>' + docLabel + ' เลขที่ ' + escapeHtml_(found.obj.RunningNo) + ' — อนุมัติแล้ว</h3>' +
      '<p>เรียน ฝ่ายบัญชี และ GA เพื่อทราบ — เอกสารนี้ได้รับการอนุมัติเรียบร้อยแล้ว</p>' +
      '<p><b>หน่วยงาน:</b> ' + escapeHtml_(found.obj.FromDept || '-') + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
      '<tr style="background:#f0f4f8;"><th style="border:1px solid #ddd;padding:6px;">#</th><th style="border:1px solid #ddd;padding:6px;">รหัส</th><th style="border:1px solid #ddd;padding:6px;">รายการ</th></tr>' +
      itemsHtml +
      '</table>' +
      '</div>';
    MailApp.sendEmail({ to: recipients.join(','), subject: '[อนุมัติแล้ว] ' + docLabel + ' ' + found.obj.RunningNo + ' — ' + CONFIG.COMPANY_NAME, htmlBody: html });

    // บันทึกเวลาที่ส่งแจ้งเตือนไว้ในชีตเอกสาร ให้หน้ารายการแสดงสถานะ "ส่งแล้ว/ยังไม่ส่ง" ได้
    // (guard idx.NotifiedAt !== undefined เผื่อยังไม่ได้รัน setup() ใหม่เพื่อเพิ่มคอลัมน์นี้)
    if (found.idx.NotifiedAt !== undefined) {
      const sh = getSS_().getSheetByName(sheetName);
      sh.getRange(found.rowNum, found.idx.NotifiedAt + 1).setValue(new Date());
    }

    logActivity_(id, 'NOTIFY_ACCOUNTING_GA', body.by || '', 'แจ้งเตือนฝ่ายบัญชี/GA สำหรับ' + docLabel + ' ' + found.obj.RunningNo);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
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
      // ราคาซาก = มูลค่าตามบัญชีเสมอ
      if (idx.ScrapPrice !== undefined) newRow[idx.ScrapPrice] = newRow[idx.BookValue];
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

  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่' };
  if (!canManageDept_(user, body.fromDept)) return { ok: false, error: 'ไม่มีสิทธิ์สร้างใบโอนย้ายให้หน่วยงานนี้' };

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
    exportDocToSharePointSafe_('transfer', getTransferFull_(transferId));
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
    exportDocToSharePointSafe_('transfer', getTransferFull_(transferId));
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

  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่' };
  if (!canManageDept_(user, body.fromDept)) return { ok: false, error: 'ไม่มีสิทธิ์สร้างใบขายออกให้หน่วยงานนี้' };

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
    exportDocToSharePointSafe_('sale', getSaleFull_(saleId));
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
    exportDocToSharePointSafe_('sale', getSaleFull_(saleId));
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

  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่' };
  if (!canManageDept_(user, body.fromDept)) return { ok: false, error: 'ไม่มีสิทธิ์สร้างใบตัดชำรุดให้หน่วยงานนี้' };

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
    exportDocToSharePointSafe_('writeoff', getWriteOffFull_(writeOffId));
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
    exportDocToSharePointSafe_('writeoff', getWriteOffFull_(writeOffId));
  }

  logActivity_(writeOffId, 'WRITEOFF_' + decision.toUpperCase(), found.obj.ApproverName || found.obj.ApproverEmail, body.comment || '');

  sendWriteOffDecisionNotification_(found.obj, decision, body.comment || '');

  return { ok: true, data: { writeOffId, status: decision } };
}

// ============================================================
// IMAGE UPLOAD (Google Drive)
// ============================================================
function uploadImage_(body) {
  const user = getRequestingUser_(body.password);
  if (!user) return { ok: false, error: 'กรุณาเข้าสู่ระบบก่อนอัปโหลดรูปภาพ' };
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
// SHAREPOINT EXPORT — บันทึก PDF ใบโอนย้าย/ขายออก/ตัดชำรุด ที่อนุมัติแล้วขึ้น SharePoint อัตโนมัติ
// ============================================================
// ต้องตั้งค่า Script Properties 3 ค่านี้ก่อนใช้งาน (Apps Script editor > Project Settings > Script Properties):
//   SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET
// และต้องเปิด Advanced Service "Drive API" (Services > + > Drive API) เพื่อแปลง HTML เป็น PDF
// ถ้ายังไม่ได้ตั้งค่า ฟังก์ชันนี้จะข้ามไปเงียบๆ ไม่ทำให้การสร้าง/อนุมัติเอกสารล้มเหลว

function getSharePointCreds_() {
  const props = PropertiesService.getScriptProperties();
  const tenantId = props.getProperty('SHAREPOINT_TENANT_ID');
  const clientId = props.getProperty('SHAREPOINT_CLIENT_ID');
  const clientSecret = props.getProperty('SHAREPOINT_CLIENT_SECRET');
  if (!tenantId || !clientId || !clientSecret) return null;
  return { tenantId, clientId, clientSecret };
}

function getSharePointAccessToken_() {
  const creds = getSharePointCreds_();
  if (!creds) return null;

  const cache = CacheService.getScriptCache();
  const cached = cache.get('sp_access_token');
  if (cached) return cached;

  const tokenUrl = 'https://login.microsoftonline.com/' + creds.tenantId + '/oauth2/v2.0/token';
  const resp = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    payload: {
      client_id: creds.clientId,
      client_secret: creds.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    },
    muteHttpExceptions: true
  });
  const data = JSON.parse(resp.getContentText());
  if (!data.access_token) throw new Error('ขอ access token จาก Microsoft ไม่สำเร็จ: ' + resp.getContentText());
  cache.put('sp_access_token', data.access_token, Math.min((data.expires_in || 3600) - 60, 1500));
  return data.access_token;
}

// หา driveId ของ document library ปลายทางบน SharePoint (แคชไว้ใน Script Properties เพราะแทบไม่เปลี่ยน)
function getSharePointDriveId_(token) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('SHAREPOINT_DRIVE_ID');
  if (cached) return cached;

  const siteResp = UrlFetchApp.fetch(
    'https://graph.microsoft.com/v1.0/sites/' + CONFIG.SHAREPOINT_HOSTNAME + ':' + CONFIG.SHAREPOINT_SITE_PATH,
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const siteData = JSON.parse(siteResp.getContentText());
  if (!siteData.id) throw new Error('ไม่พบไซต์ SharePoint (' + CONFIG.SHAREPOINT_SITE_PATH + '): ' + siteResp.getContentText());

  const drivesResp = UrlFetchApp.fetch(
    'https://graph.microsoft.com/v1.0/sites/' + siteData.id + '/drives',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const drivesData = JSON.parse(drivesResp.getContentText());
  const drive = (drivesData.value || []).find(d => d.name === 'Documents' || d.name === 'Shared Documents') || (drivesData.value || [])[0];
  if (!drive) throw new Error('ไม่พบ document library บนไซต์นี้: ' + drivesResp.getContentText());

  props.setProperty('SHAREPOINT_DRIVE_ID', drive.id);
  return drive.id;
}

// แปลง HTML เป็น PDF โดยอาศัย Google Drive แปลงไฟล์ให้ (ต้องเปิด Advanced Service "Drive API" ก่อน)
// รองรับทั้ง Drive API v2 (Files.insert/title) และ v3 (Files.create/name) เพราะ Apps Script
// ผูก Advanced Service เวอร์ชันไหนให้ก็ได้แล้วแต่ตอนเพิ่ม ไม่ควรอิงว่าเป็นเวอร์ชันใดเวอร์ชันหนึ่งตายตัว
function htmlToPdfBlob_(html, fileName) {
  const htmlBlob = Utilities.newBlob(html, MimeType.HTML, fileName + '.html');
  let file;
  if (Drive.Files.create) {
    // Drive API v3
    file = Drive.Files.create({ name: fileName, mimeType: MimeType.GOOGLE_DOCS }, htmlBlob);
  } else {
    // Drive API v2
    file = Drive.Files.insert({ title: fileName, mimeType: MimeType.GOOGLE_DOCS }, htmlBlob, { convert: true });
  }
  try {
    return DriveApp.getFileById(file.id).getAs(MimeType.PDF);
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true); // ลบไฟล์ Google Doc ชั่วคราวทิ้ง เหลือแค่ PDF ที่อัปโหลดต่อ
  }
}

function uploadBlobToSharePoint_(blob, fileName, subfolder) {
  const token = getSharePointAccessToken_();
  if (!token) return { ok: false, error: 'ยังไม่ได้ตั้งค่า SharePoint (Script Properties)' };
  const driveId = getSharePointDriveId_(token);
  const path = CONFIG.SHAREPOINT_FOLDER + '/' + subfolder + '/' + fileName;
  const encodedPath = path.split('/').map(encodeURIComponent).join('/');
  const resp = UrlFetchApp.fetch(
    'https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:/' + encodedPath + ':/content',
    {
      method: 'put',
      contentType: blob.getContentType(),
      payload: blob.getBytes(),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() >= 300) return { ok: false, error: 'อัปโหลด SharePoint ไม่สำเร็จ: ' + resp.getContentText() };
  return { ok: true };
}

function fmtDateServer_(d) {
  if (!d) return '-';
  const dt = new Date(d);
  if (isNaN(dt)) return String(d);
  const buddhistYear = dt.getFullYear() + 543;
  return ('0' + dt.getDate()).slice(-2) + '/' + ('0' + (dt.getMonth() + 1)).slice(-2) + '/' + buddhistYear;
}

// สไตล์นี้ตั้งใจให้ตรงกับ .doc/.doc-table/.doc-sign ฯลฯ ในหน้าพิมพ์ของ index.html (buildDocHtml/buildSaleDocHtml/
// buildWriteOffDocHtml) เพื่อให้ PDF ที่อัปโหลดขึ้น SharePoint หน้าตาเหมือนกับที่ผู้ใช้เห็นตอนกด "พิมพ์" ในแอปเป๊ะๆ
// ใช้ค่าสีจริงแทนตัวแปร CSS (--navy ฯลฯ) เพราะตัวแปลง HTML→PDF ของ Drive ไม่รองรับ CSS custom properties
function pdfDocStyle_() {
  return '<style>' +
    'body{font-family:"Sarabun","Angsana New",sans-serif;font-size:14px;line-height:1.6;color:#222;padding:28px 34px;}' +
    '.doc-header{text-align:center;font-weight:700;font-size:18px;margin-bottom:2px;}' +
    '.doc-title{text-align:center;font-weight:700;font-size:17px;margin:10px 0 16px;text-decoration:underline;}' +
    '.doc-meta{text-align:right;font-size:13px;margin-bottom:10px;}' +
    '.doc-field{margin:4px 0;font-size:13.5px;}' +
    '.doc-field b{display:inline-block;min-width:70px;}' +
    'table.doc-table{width:100%;border-collapse:collapse;margin-top:14px;font-size:12.5px;table-layout:fixed;}' +
    'table.doc-table th{border:1px solid #dde3ec;padding:8px;background:#eaf0f8;color:#1a3c6e;font-weight:600;text-align:center;}' +
    'table.doc-table td{border:1px solid #dde3ec;padding:8px;text-align:center;vertical-align:middle;word-wrap:break-word;overflow-wrap:break-word;}' +
    'table.doc-table td.left{text-align:left;}' +
    'table.doc-table td.money{text-align:right;}' +
    '.doc-img-gallery{text-align:center;}' +
    '.doc-img-gallery img{width:26px;height:26px;object-fit:cover;border-radius:3px;border:1px solid #ddd;margin:1px;}' +
    '.doc-note{font-size:11.5px;color:#555;margin-top:16px;border-top:1px dashed #ccc;padding-top:10px;}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;}' +
    '</style>';
}

function pdfImgGallery_(images) {
  const list = (images || []).filter(Boolean);
  if (!list.length) return '-';
  return '<div class="doc-img-gallery">' + list.map(u => '<img src="' + escapeHtml_(u) + '">').join('') + '</div>';
}

function pdfStatusBadge_(status, isSale) {
  const labels = { Draft: 'ฉบับร่าง', PendingApproval: 'รออนุมัติ', Approved: (isSale ? 'ขายแล้ว' : 'อนุมัติแล้ว'), Rejected: 'ไม่อนุมัติ', Voided: 'ยกเลิกแล้ว' };
  const colors = { Draft: ['#e2e3e5', '#555'], PendingApproval: ['#fff3cd', '#b8860b'], Approved: ['#d4edda', '#1a7d3c'], Rejected: ['#f8d7da', '#c0392b'], Voided: ['#e2e3e5', '#555'] };
  const c = colors[status] || colors.Draft;
  return '<span class="badge" style="background:' + c[0] + ';color:' + c[1] + ';">' + escapeHtml_(labels[status] || status) + '</span>';
}

// ใช้ตารางไม่มีเส้นขอบแทน .doc-sign แบบ flexbox เดิม เพราะ Google Docs (ตัวแปลง HTML เป็น PDF)
// ไม่รองรับ display:flex ดีนัก ทำให้เส้นขีดกับชื่อ/วันที่ใต้เส้นจัดกึ่งกลางไม่ตรงกัน — ตารางแปลงได้เสถียรกว่ามาก
function pdfSignBlock_(leftLabel, leftName, rightLabel, rightName) {
  const cell = (label, name) =>
    '<td style="width:45%;text-align:center;border:none;padding:0;">' +
    '__________________<br>' +
    '(' + escapeHtml_(label) + ') ' + escapeHtml_(name || '') + '<br>' +
    'วันที่ ..........................' +
    '</td>';
  return '<table style="width:100%;border:none;margin-top:30px;font-size:13px;"><tr>' +
    cell(leftLabel, leftName) +
    '<td style="width:10%;border:none;"></td>' +
    cell(rightLabel, rightName) +
    '</tr></table>';
}

function pdfApprovalFooter_(obj, isSale) {
  if (obj.Status === 'Draft') return '';
  let html = '<div style="margin-top:20px;padding-top:14px;border-top:1px solid #ddd;font-size:12.5px;">' +
    '<b>สถานะการอนุมัติ:</b> ' + pdfStatusBadge_(obj.Status, isSale) + '&nbsp; ';
  if (obj.ApproverName || obj.ApproverEmail) html += 'โดย ' + escapeHtml_(obj.ApproverName || obj.ApproverEmail);
  if (obj.ApprovedAt) html += ' เมื่อ ' + fmtDateServer_(obj.ApprovedAt);
  if (obj.ApproverComment) html += '<br><b>ความเห็น:</b> ' + escapeHtml_(obj.ApproverComment);
  html += '</div>';
  return html;
}

function buildTransferPdfHtml_(t) {
  const rows = (t.Items || []).map((it, i) =>
    '<tr><td>' + (i + 1) + '</td><td>' + pdfImgGallery_(it.Images) + '</td><td>' + escapeHtml_(it.AssetID) + '</td><td class="left">' + escapeHtml_(it.AssetName) +
    '</td><td>' + escapeHtml_(it.FromDeptName) + '<br><span style="font-size:11px;color:#666;">' + escapeHtml_(it.FromSignName) + '</span></td>' +
    '<td>' + escapeHtml_(it.ToDeptName) + '<br><span style="font-size:11px;color:#666;">' + escapeHtml_(it.ToSignName) + '</span></td>' +
    '<td>' + escapeHtml_(it.Remark) + '</td></tr>'
  ).join('');
  const firstItem = (t.Items && t.Items[0]) || {};
  return '<html><head><meta charset="UTF-8">' + pdfDocStyle_() + '</head><body>' +
    '<div class="doc-header">' + escapeHtml_(CONFIG.COMPANY_NAME) + '</div>' +
    '<div class="doc-title">ใบโอนย้ายทรัพย์สิน</div>' +
    '<div class="doc-meta"><b>เลขที่</b> ' + escapeHtml_(t.RunningNo) + ' &nbsp;&nbsp;&nbsp; <b>วันที่</b> ' + fmtDateServer_(t.CreatedAt) + '</div>' +
    '<div class="doc-field"><b>เรียน</b> ผช.ผู้จัดการส่วนบัญชีและการเงิน ทราบ</div>' +
    '<div class="doc-field"><b>เรื่อง</b> ' + (t.Subject === 'อื่นๆ' ? 'อื่นๆ (' + escapeHtml_(t.SubjectOther) + ')' : 'โอนย้าย') + '</div>' +
    '<div class="doc-field"><b>เพื่อ</b> ' + escapeHtml_(t.Purpose || '-') + '</div>' +
    '<div class="doc-field">ดังมีรายละเอียดดังนี้ &nbsp; <b>จาก</b> ' + escapeHtml_(t.FromDept) + ' &nbsp; <b>ไปยัง</b> ' + escapeHtml_(t.ToDept) + '</div>' +
    '<table class="doc-table"><tr><th style="width:5%;">ลำดับ</th><th style="width:15%;">รูปภาพ</th><th style="width:10%;">รหัส</th><th style="width:25%;">รายการ</th>' +
    '<th style="width:15%;">ผู้โอน<br>หน่วยงาน/ผู้ลงชื่อ</th><th style="width:15%;">ผู้รับโอน<br>หน่วยงาน/ผู้ลงชื่อ</th><th style="width:15%;">หมายเหตุ</th></tr>' + rows + '</table>' +
    '<div class="doc-note">หมายเหตุ : กรุณากำหนดเลขที่ running number ดังนี้ xxx/001/yy<br>' +
    'xxx หมายถึง หน่วยงาน (ดูรหัสได้ที่เมนู "ตั้งค่ารหัสหน่วยงาน") &nbsp; 001 หมายถึง ลำดับเลขที่เอกสาร &nbsp; yy หมายถึง ปี พ.ศ. 2 หลัก</div>' +
    pdfSignBlock_('ผู้โอน', firstItem.FromSignName, 'ผู้รับโอน', firstItem.ToSignName) +
    '<div style="margin-top:26px;font-size:13px;"><b>รับทราบโดย</b><br>' +
    '1. ผจก.ฝ่าย/ผจก.ส่วน (โอน) &nbsp;&nbsp; วันที่ <span style="white-space:nowrap;">..........................</span><br>' +
    '2. ผจก.ฝ่าย/ผจก.ส่วน (รับโอน) &nbsp;&nbsp; วันที่ <span style="white-space:nowrap;">..........................</span></div>' +
    pdfApprovalFooter_(t, false) +
    '</body></html>';
}

function buildSalePdfHtml_(s) {
  const rows = (s.Items || []).map((it, i) =>
    '<tr><td>' + (i + 1) + '</td><td>' + pdfImgGallery_(it.Images) + '</td><td>' + escapeHtml_(it.AssetID) + '</td><td class="left">' + escapeHtml_(it.AssetName) +
    '</td><td class="money">' + fmtMoneyServer_(it.ScrapPrice) + '</td><td class="money">' + fmtMoneyServer_(it.AuctionPrice) +
    '</td><td class="money">' + fmtMoneyServer_(it.SalePrice) + '</td><td>' + escapeHtml_(it.Remark) + '</td></tr>'
  ).join('');
  const total = (s.Items || []).reduce((sum, it) => sum + (parseFloat(it.SalePrice) || 0), 0);
  return '<html><head><meta charset="UTF-8">' + pdfDocStyle_() + '</head><body>' +
    '<div class="doc-header">' + escapeHtml_(CONFIG.COMPANY_NAME) + '</div>' +
    '<div class="doc-title">ใบขายออกทรัพย์สิน</div>' +
    '<div class="doc-meta"><b>เลขที่</b> ' + escapeHtml_(s.RunningNo) + ' &nbsp;&nbsp;&nbsp; <b>วันที่</b> ' + fmtDateServer_(s.CreatedAt) + '</div>' +
    '<div class="doc-field"><b>หน่วยงาน</b> ' + escapeHtml_(s.FromDept) + '</div>' +
    '<div class="doc-field"><b>ผู้ซื้อ/ผู้ประมูลได้</b> ' + escapeHtml_(s.Buyer || '-') + '</div>' +
    '<div class="doc-field"><b>หมายเหตุ</b> ' + escapeHtml_(s.Remark || '-') + '</div>' +
    '<table class="doc-table"><tr><th style="width:5%;">ลำดับ</th><th style="width:15%;">รูปภาพ</th><th style="width:10%;">รหัส</th><th style="width:21%;">รายการ</th>' +
    '<th style="width:12%;">ราคาซาก</th><th style="width:12%;">ราคาประมูล</th><th style="width:12%;">ราคาขาย</th><th style="width:13%;">หมายเหตุ</th></tr>' + rows +
    '<tr><td colspan="6" style="text-align:right;font-weight:600;">รวมราคาขาย</td><td style="font-weight:700;">' + fmtMoneyServer_(total) + '</td><td></td></tr></table>' +
    pdfSignBlock_('ผู้บันทึก', s.CreatedBy, 'ผู้อนุมัติ', s.ApproverName) +
    pdfApprovalFooter_(s, true) +
    '</body></html>';
}

function buildWriteOffPdfHtml_(w) {
  const rows = (w.Items || []).map((it, i) =>
    '<tr><td>' + (i + 1) + '</td><td>' + pdfImgGallery_(it.Images) + '</td><td>' + escapeHtml_(it.AssetID) + '</td><td class="left">' + escapeHtml_(it.AssetName) +
    '</td><td class="money">' + fmtMoneyServer_(it.ScrapPrice) + '</td><td>' + escapeHtml_(it.Remark) + '</td></tr>'
  ).join('');
  const total = (w.Items || []).reduce((sum, it) => sum + (parseFloat(it.ScrapPrice) || 0), 0);
  return '<html><head><meta charset="UTF-8">' + pdfDocStyle_() + '</head><body>' +
    '<div class="doc-header">' + escapeHtml_(CONFIG.COMPANY_NAME) + '</div>' +
    '<div class="doc-title">ใบตัดชำรุดทรัพย์สิน</div>' +
    '<div class="doc-meta"><b>เลขที่</b> ' + escapeHtml_(w.RunningNo) + ' &nbsp;&nbsp;&nbsp; <b>วันที่</b> ' + fmtDateServer_(w.CreatedAt) + '</div>' +
    '<div class="doc-field"><b>หน่วยงาน</b> ' + escapeHtml_(w.FromDept) + '</div>' +
    '<div class="doc-field"><b>สาเหตุชำรุด</b> ' + escapeHtml_(w.Reason || '-') + '</div>' +
    '<div class="doc-field"><b>หมายเหตุ</b> ' + escapeHtml_(w.Remark || '-') + '</div>' +
    '<table class="doc-table"><tr><th style="width:6%;">ลำดับ</th><th style="width:16%;">รูปภาพ</th><th style="width:12%;">รหัส</th><th style="width:30%;">รายการ</th>' +
    '<th style="width:16%;">ราคาซาก</th><th style="width:20%;">หมายเหตุ</th></tr>' + rows +
    '<tr><td colspan="4" style="text-align:right;font-weight:600;">รวมราคาซาก</td><td style="font-weight:700;">' + fmtMoneyServer_(total) + '</td><td></td></tr></table>' +
    pdfSignBlock_('ผู้บันทึก', w.CreatedBy, 'ผู้อนุมัติ', w.ApproverName) +
    pdfApprovalFooter_(w, false) +
    '</body></html>';
}

// เรียกจาก createTransfer_/createSale_/createWriteOff_ (กรณีอนุมัติอัตโนมัติ) และ decideTransfer_/decideSale_/decideWriteOff_
// (กรณีอนุมัติผ่านลิงก์อีเมล) เพื่อบันทึก PDF ของเอกสารที่อนุมัติเสร็จสมบูรณ์แล้วขึ้น SharePoint
// ห่อด้วย try/catch เสมอ — ถ้า SharePoint ล่มหรือยังไม่ได้ตั้งค่า ต้องไม่ทำให้การสร้าง/อนุมัติเอกสารหลักล้มเหลว
function exportDocToSharePointSafe_(kind, fullObj) {
  try {
    if (!getSharePointCreds_()) return; // ยังไม่ได้ตั้งค่า ข้ามไปเงียบๆ
    let html, subfolder, titlePrefix;
    if (kind === 'transfer') { html = buildTransferPdfHtml_(fullObj); subfolder = 'โอนย้าย'; titlePrefix = 'ใบโอนย้าย'; }
    else if (kind === 'sale') { html = buildSalePdfHtml_(fullObj); subfolder = 'ขายออก'; titlePrefix = 'ใบขายออก'; }
    else if (kind === 'writeoff') { html = buildWriteOffPdfHtml_(fullObj); subfolder = 'ตัดชำรุด'; titlePrefix = 'ใบตัดชำรุด'; }
    else return;

    const safeRunningNo = String(fullObj.RunningNo || '').replace(/\//g, '-');
    const fileName = titlePrefix + '_' + safeRunningNo + '.pdf';
    const pdfBlob = htmlToPdfBlob_(html, titlePrefix + '_' + safeRunningNo);
    const result = uploadBlobToSharePoint_(pdfBlob, fileName, subfolder);
    if (!result.ok) {
      logActivity_('', 'SHAREPOINT_EXPORT_FAILED', 'system', (fullObj.RunningNo || '') + ': ' + result.error);
    }
  } catch (err) {
    logActivity_('', 'SHAREPOINT_EXPORT_FAILED', 'system', (fullObj && fullObj.RunningNo || '') + ': ' + String(err));
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
