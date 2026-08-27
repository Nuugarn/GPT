/*
  ====== วิธีใช้งาน ======
  Sheet "Transactions" หัวตาราง (แถวที่ 1):
  id | datetime | location | type | category | title | detail | amount | imageUrl

  Sheet "Categories" หัวตาราง (แถวที่ 1):
  name

  IMPORTANT:
  - คอลัมน์ datetime เป็น Plain Text เพื่อไม่ให้ timezone ถูกเลื่อน
  - location ใช้แยกรายรับ/รายจ่ายของแต่ละสถานที่
*/

const SPREADSHEET_ID = "1WSPdRTWkkOCDzdvOFJuWfWOfd1LGIrqpkSQBLlOq6YE";
const DRIVE_FOLDER_ID = "1IWhPKf0x0_hX9Rj5ryZExAMkihqpNSia";
const TRANSACTIONS_SHEET = "Transactions";
const CATEGORIES_SHEET = "Categories";
const LOCATIONS_SHEET = "Locations";
const BANGKOK_TIMEZONE = "Asia/Bangkok";
// บ้านเริ่มต้นของระบบ; รายชื่อจริงสามารถจัดการต่อได้จาก Sheet Locations
const DEFAULT_LOCATIONS = ["บ้าน 1", "บ้าน 2"];

function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Keep all displayed/parsed Date values aligned to Bangkok time.
  try { ss.setSpreadsheetTimeZone(BANGKOK_TIMEZONE); } catch (ignore) {}
  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const txValues = txSheet.getDataRange().getValues();
  const txHeaders = txValues.shift();

  const transactions = txValues
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      txHeaders.forEach((h, i) => obj[h] = row[i]);

      obj.datetime = normalizeDatetime(obj.datetime);
      obj.location = normalizeLocation(obj.location);
      return obj;
    });

  const catSheet = ss.getSheetByName(CATEGORIES_SHEET);
  const catValues = catSheet ? catSheet.getDataRange().getValues() : [];
  if (catValues.length) catValues.shift();
  const categories = catValues.map(row => String(row[0] || "").trim()).filter(Boolean);

  const locations = getLocations(ss, transactions);
  return jsonResponse({ transactions, categories, locations });
}

function getLocations(ss, transactions) {
  let sheet = ss.getSheetByName(LOCATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOCATIONS_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([["name", "active"]]);
  } else if (sheet.getLastColumn() < 2) {
    sheet.getRange(1, 2).setValue("active");
  }

  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const existing = rows.map(row => normalizeLocation(row[0])).filter(Boolean);
  const additions = DEFAULT_LOCATIONS.filter(name => !existing.includes(name));
  if (additions.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, additions.length, 2).setValues(additions.map(name => [name, true]));
  }

  const refreshedLast = sheet.getLastRow();
  const configuredRows = refreshedLast > 1 ? sheet.getRange(2, 1, refreshedLast - 1, 2).getValues() : [];
  const configured = configuredRows
    .filter(row => row[0] && row[1] !== false && String(row[1]).toLowerCase() !== "false")
    .map(row => normalizeLocation(row[0]))
    .filter(Boolean);

  // Keep legacy locations visible without silently assigning them to a house.
  const fromTransactions = transactions.map(t => t.location).filter(Boolean);
  return [...new Set([...DEFAULT_LOCATIONS, ...configured, ...fromTransactions])].sort((a, b) => {
    const ai = DEFAULT_LOCATIONS.indexOf(a);
    const bi = DEFAULT_LOCATIONS.indexOf(b);
    if (ai !== -1 || bi !== -1) return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
    return a.localeCompare(b, "th");
  });
}

function normalizeLocation(value) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";
  const compact = text.replace(/\s+/g, "");
  if (compact === "บ้าน1") return "บ้าน 1";
  if (compact === "บ้าน2") return "บ้าน 2";
  return text;
}

function normalizeDatetime(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, BANGKOK_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  }

  let text = String(value == null ? "" : value).trim();
  if (!text) return "";
  text = text.replace(/\s+/, "T");
  const match = text.match(/^(\d{4}-\d{2}-\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return text;
  return `${match[1]}T${match[2]}:${match[3]}:${match[4] || "00"}`;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");

    // ไม่รับ Secret จาก Frontend อีกต่อไป เพราะค่าที่ฝังใน HTML ไม่ใช่ความลับจริง
    if (body.action === "setupHouseSystem") {
      return setupHouseSystem();
    }

    if (body.action === "addTransaction") {
      return addTransaction(body);
    }

    if (body.action === "addCategory") {
      return addCategory(body);
    }

    return jsonResponse({ ok: false, error: "unknown action" });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function setupHouseSystem() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.setSpreadsheetTimeZone(BANGKOK_TIMEZONE);
  getLocations(ss, []);
  return jsonResponse({ ok: true, locations: getLocations(ss, []) });
}

// Legacy rows with blank location are intentionally preserved.
// Assign them only when the owner knows the correct house.
function assignTransactionLocation(id, location) {
  const normalized = normalizeLocation(location);
  if (!normalized) throw new Error("missing location");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idCol = headers.indexOf("id");
  const locCol = headers.indexOf("location");
  if (idCol < 0 || locCol < 0) throw new Error("Transactions headers are invalid");
  const rowIndex = values.findIndex(row => String(row[idCol]) === String(id));
  if (rowIndex < 0) throw new Error("transaction not found");
  sheet.getRange(rowIndex + 2, locCol + 1).setNumberFormat("@").setValue(normalized);
  return jsonResponse({ ok: true, id: id, location: normalized });
}

function addTransaction(body) {
  if (!body.type || !["income", "expense"].includes(body.type)) {
    return jsonResponse({ ok: false, error: "invalid type" });
  }
  if (!String(body.title || "").trim()) {
    return jsonResponse({ ok: false, error: "missing title" });
  }
  if (!body.datetime) {
    return jsonResponse({ ok: false, error: "missing datetime" });
  }

  const location = normalizeLocation(body.location);
  if (!location) {
    return jsonResponse({ ok: false, error: "missing location" });
  }

  const amount = Number(body.amount);
  if (!Number.isFinite(amount) || amount <= 0 || amount > 1000000000) {
    return jsonResponse({ ok: false, error: "invalid amount" });
  }

  const datetime = normalizeDatetime(body.datetime);
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(datetime)) {
    return jsonResponse({ ok: false, error: "invalid datetime" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  let imageUrl = "";

  if (body.imageBase64) {
    if (String(body.imageBase64).length > 7500000) {
      return jsonResponse({ ok: false, error: "image too large" });
    }
    imageUrl = saveImageToDrive(body.imageBase64, body.imageName || "slip.jpg");
  }

  // เก็บเวลาที่ผู้ใช้กรอกเป็น local time ของไทยตรง ๆ
  const id = Date.now();
  const row = sheet.getLastRow() + 1;

  // คอลัมน์ B = datetime และ C = location
  sheet.getRange(row, 2).setNumberFormat("@");
  sheet.getRange(row, 3).setNumberFormat("@");

  sheet.getRange(row, 1, 1, 9).setValues([[
    id,
    datetime,
    location,
    body.type,
    body.category || "",
    body.title,
    body.detail || "",
    amount,
    imageUrl
  ]]);

  return jsonResponse({ ok: true, id, imageUrl, location });
}

function addCategory(body) {
  if (!body.name) {
    return jsonResponse({ ok: false, error: "missing category name" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CATEGORIES_SHEET);
  const existing = sheet.getDataRange().getValues().flat();

  if (existing.includes(body.name)) {
    return jsonResponse({ ok: true, alreadyExists: true });
  }

  sheet.appendRow([body.name]);
  return jsonResponse({ ok: true });
}

function saveImageToDrive(base64Data, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const matches = base64Data.match(/^data:(.+);base64,(.+)$/);
  const mimeType = matches ? matches[1] : "image/jpeg";
  const data = matches ? matches[2] : base64Data;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data),
    mimeType,
    fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
