const SPREADSHEET_ID = "12-5YwJTdQ4G73QzLM6kVgxAcJ9gZnwzcpU20Xdk1gIs";
const SHEET_NAME = "Reservasi";
const TIMEZONE = "Asia/Makassar";

// STRUKTUR FINAL A:K
const HEADERS = ["Kode","Nama","WhatsApp","Layanan","Barber","Tanggal","Jam","Total","Catatan","Status","Dibuat"];
const STATUS_OPTIONS = ["Menunggu Konfirmasi","Dikonfirmasi","Selesai","Dibatalkan"];
const BARBERS = { bj: "Bang Jaka", kr: "Kang Rudi" };
// Kompatibel dengan frontend yang mengirim b1/b2.
const BARBER_ALIASES = { b1: "bj", b2: "kr", bj: "bj", kr: "kr" };

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = safeString_(p.action).toLowerCase();
    if (action === "slots") return jsonOutput_(getSlots_(p.barberId, p.date));
    if (action === "lookup") return jsonOutput_(lookupBookingByCode_(p.code));
    if (action === "health") return jsonOutput_({success:true,message:"API Mannuruki Barbershop aktif.",endpoints:["health","slots","lookup"],sheet:SHEET_NAME});
    return jsonOutput_({success:true,message:"API Mannuruki Barbershop aktif.",endpoints:["health","slots","lookup"]});
  } catch (err) {
    return jsonOutput_({success:false,message:getErrorMessage_(err)});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const data = parsePostData_(e);
    validateBooking_(data);
    const sheet = getSheet_();
    ensureSchema_(sheet);

    if (codeExists_(sheet, data.code)) {
      return jsonOutput_({success:true,duplicate:true,message:"Reservasi dengan kode tersebut sudah tersimpan.",code:data.code});
    }

    const booked = getBookedSlots_(sheet, data.barberId, data.date);
    if (booked.indexOf(data.time) !== -1) {
      return jsonOutput_({success:false,slotTaken:true,message:"Jam tersebut sudah dipesan. Silakan pilih jam lain."});
    }

    const services = data.services.map(function(s){return safeString_(s);}).filter(Boolean).join(", ");
    const now = new Date();
    const row = [data.code,data.name,data.phone,services,data.barberName,data.date,data.time,Number(data.total),data.notes,"Menunggu Konfirmasi",now];
    sheet.appendRow(row);
    SpreadsheetApp.flush();

    const rowNumber = sheet.getLastRow();
    sheet.getRange(rowNumber,8).setNumberFormat('"Rp" #,##0');
    sheet.getRange(rowNumber,11).setNumberFormat("dd/mm/yyyy hh:mm:ss");

    return jsonOutput_({success:true,message:"Reservasi berhasil disimpan ke Google Sheets.",code:data.code,booking:makeBookingObject_(row)});
  } catch (err) {
    return jsonOutput_({success:false,message:getErrorMessage_(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ===================== SHEET & MIGRASI ===================== */
function setupSheet() {
  const sheet = getSheet_();
  ensureSchema_(sheet);
  configureSheet_(sheet);
  return "Sheet Reservasi berhasil disiapkan. Struktur final A:K.";
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureSchema_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const header = sheet.getRange(1,1,1,lastColumn).getValues()[0].map(safeString_);

  const legacy = ["Kode","Nama","WhatsApp","Layanan","Barber","Barber ID","Tanggal","Jam","Total","Catatan","Status","Dibuat"];
  const finalOk = header.slice(0,HEADERS.length).join("|") === HEADERS.join("|");
  const legacyOk = header.slice(0,legacy.length).join("|") === legacy.join("|");

  if (legacyOk) {
    migrateLegacy12To11_(sheet);
  } else if (!finalOk) {
    const empty = header.every(function(v){return !v;});
    if (empty && lastRow <= 1) {
      sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    } else if (lastRow <= 1) {
      sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    } else {
      throw new Error("Struktur sheet tidak dikenali. Gunakan format A:K Mannuruki Barbershop.");
    }
  }
  configureSheet_(sheet);
}

function migrateLegacy12To11_(sheet) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  if (lastRow < 2) {
    if (sheet.getMaxColumns() >= 12) sheet.deleteColumn(12);
    return;
  }

  // Lama: A Kode, B Nama, C WhatsApp, D Layanan, E Barber,
  // F Barber ID, G Tanggal, H Jam, I Total, J Catatan, K Status, L Dibuat.
  const oldRows = sheet.getRange(2,1,lastRow-1,12).getValues();
  const newRows = oldRows.map(function(r){
    const barberId = normalizeBarberId_(r[5]);
    const barberName = safeString_(r[4]) || BARBERS[barberId] || "";
    return [
      safeString_(r[0]).toUpperCase(), safeString_(r[1]), normalizePhone_(r[2]), safeString_(r[3]),
      barberName, normalizeSheetDate_(r[6]), normalizeSheetTime_(r[7]), normalizeNumber_(r[8]),
      safeString_(r[9]), normalizeStatus_(r[10]), normalizeCreatedAt_(r[11])
    ];
  });
  sheet.getRange(2,1,newRows.length,HEADERS.length).setValues(newRows);
  if (sheet.getMaxColumns() >= 12) sheet.deleteColumn(12);
  SpreadsheetApp.flush();
}

function configureSheet_(sheet) {
  sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("F:G").setNumberFormat("@");
  sheet.getRange("H:H").setNumberFormat('"Rp" #,##0');
  sheet.getRange("K:K").setNumberFormat("dd/mm/yyyy hh:mm:ss");

  // VALIDASI STATUS HANYA DI J2:J, bukan K.
  const rows = Math.max(sheet.getMaxRows()-1,1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS,true)
    .setAllowInvalid(false)
    .setHelpText("Pilih status reservasi.")
    .build();
  sheet.getRange(2,10,rows,1).setDataValidation(rule);

  const widths = [130,180,140,300,140,130,90,130,250,180,170];
  widths.forEach(function(w,i){sheet.setColumnWidth(i+1,w);});
}

function normalizeStatus_(value) {
  const s = safeString_(value).toLowerCase();
  if (!s || s === "menunggu" || s === "menunggu konfirmasi") return "Menunggu Konfirmasi";
  if (s === "dikonfirmasi") return "Dikonfirmasi";
  if (s === "selesai") return "Selesai";
  if (s === "dibatalkan" || s === "cancelled" || s === "canceled") return "Dibatalkan";
  return "Menunggu Konfirmasi";
}

/* ===================== POST VALIDATION ===================== */
function parsePostData_(e) {
  if (!e) throw new Error("Data reservasi tidak ditemukan.");
  let raw = "";
  if (e.parameter && e.parameter.data) raw = e.parameter.data;
  else if (e.postData && e.postData.contents) raw = e.postData.contents;
  if (!raw) throw new Error("Data reservasi tidak ditemukan.");
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") throw new Error();
    return parsed;
  } catch (_) { throw new Error("Format data reservasi tidak valid."); }
}

function validateBooking_(data) {
  if (!data || typeof data !== "object") throw new Error("Format data reservasi tidak valid.");
  data.code = safeString_(data.code).toUpperCase();
  if (!/^KNR-[A-Z0-9]{5,10}$/.test(data.code)) throw new Error("Kode reservasi tidak valid.");

  data.name = safeString_(data.name);
  if (data.name.length < 3 || data.name.length > 100) throw new Error("Nama pelanggan harus 3–100 karakter.");

  data.phone = normalizePhone_(data.phone);
  if (data.phone.length < 9 || data.phone.length > 15) throw new Error("Nomor WhatsApp tidak valid.");

  data.barberId = normalizeBarberId_(data.barberId);
  if (!BARBERS[data.barberId]) {
    const name = safeString_(data.barberName).toLowerCase();
    if (name === "bang jaka") data.barberId = "bj";
    if (name === "kang rudi") data.barberId = "kr";
  }
  if (!BARBERS[data.barberId]) throw new Error("Barber belum dipilih.");
  data.barberName = BARBERS[data.barberId];

  if (!isValidDate_(data.date)) throw new Error("Tanggal reservasi tidak valid.");
  if (!isValidTime_(data.time)) throw new Error("Jam reservasi tidak valid.");

  data.services = Array.isArray(data.services) ? data.services.map(safeString_).filter(Boolean).slice(0,10) : [];
  if (!data.services.length) throw new Error("Layanan belum dipilih.");

  data.notes = safeString_(data.notes).slice(0,500);
  data.total = Number(data.total);
  if (!Number.isFinite(data.total) || data.total < 0 || data.total > 100000000) throw new Error("Total reservasi tidak valid.");
}

/* ===================== NORMALISASI ===================== */
function normalizePhone_(value) {
  let d = safeString_(value).replace(/\D/g,"");
  if (!d) return "";
  if (d.indexOf("0") === 0) return "62" + d.slice(1);
  if (d.indexOf("62") === 0) return d;
  return d;
}

function normalizeBarberId_(id) {
  const s = safeString_(id).toLowerCase();
  return BARBER_ALIASES[s] || s;
}

function safeString_(value) { return value == null ? "" : String(value).trim(); }
function normalizeNumber_(value) { const n=Number(value); return Number.isFinite(n)?n:0; }
function normalizeCreatedAt_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return value;
  const s=safeString_(value); if(!s) return new Date();
  const d=new Date(s); return isNaN(d)?new Date():d;
}
function isCancelledStatus_(status) {
  const s=safeString_(status).toLowerCase();
  return s === "dibatalkan" || s === "cancelled" || s === "canceled";
}

/* ===================== DATE/TIME ===================== */
function isValidDate_(value) {
  const s=safeString_(value); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const p=s.split("-").map(Number), dt=new Date(p[0],p[1]-1,p[2]);
  return dt.getFullYear()===p[0] && dt.getMonth()===p[1]-1 && dt.getDate()===p[2];
}
function isValidTime_(value) {
  const s=safeString_(value); if(!/^\d{2}:\d{2}$/.test(s)) return false;
  const p=s.split(":").map(Number); return p[0]>=0 && p[0]<=23 && p[1]>=0 && p[1]<=59;
}
function normalizeSheetDate_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return Utilities.formatDate(value,TIMEZONE,"yyyy-MM-dd");
  const s=safeString_(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(m){ const d=Number(m[1]),mo=Number(m[2]),y=Number(m[3]), out=y+"-"+String(mo).padStart(2,"0")+"-"+String(d).padStart(2,"0"); if(isValidDate_(out)) return out; }
  const d=new Date(s); return isNaN(d)?s:Utilities.formatDate(d,TIMEZONE,"yyyy-MM-dd");
}
function normalizeSheetTime_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return Utilities.formatDate(value,TIMEZONE,"HH:mm");
  const s=safeString_(value), m=s.match(/^(\d{1,2}):(\d{2})/);
  return m ? String(Number(m[1])).padStart(2,"0")+":"+m[2] : s;
}
function formatDateLabel_(value) {
  const normalized=normalizeSheetDate_(value); if(!isValidDate_(normalized)) return safeString_(value);
  const p=normalized.split("-").map(Number), names=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
  return p[2]+" "+names[p[1]-1]+" "+p[0];
}

/* ===================== SLOTS ===================== */
function codeExists_(sheet,code) {
  const last=sheet.getLastRow(); if(last<2) return false;
  const target=safeString_(code).toUpperCase();
  return sheet.getRange(2,1,last-1,1).getValues().some(function(r){return safeString_(r[0]).toUpperCase()===target;});
}
function getBookedSlots_(sheet,barberId,date) {
  const last=sheet.getLastRow(); if(last<2) return [];
  const barberName=BARBERS[normalizeBarberId_(barberId)]||"";
  const rows=sheet.getRange(2,1,last-1,HEADERS.length).getValues(), out=[];
  rows.forEach(function(r){
    const storedBarber=safeString_(r[4]), storedDate=normalizeSheetDate_(r[5]), storedTime=normalizeSheetTime_(r[6]);
    if(storedBarber===barberName && storedDate===date && storedTime && !isCancelledStatus_(r[9])) out.push(storedTime);
  });
  return Array.from(new Set(out)).sort();
}
function getSlots_(barberId,date) {
  barberId=normalizeBarberId_(barberId); date=safeString_(date);
  if(!BARBERS[barberId] || !isValidDate_(date)) return {success:true,slots:[]};
  const sheet=getSheet_(); ensureSchema_(sheet);
  return {success:true,slots:getBookedSlots_(sheet,barberId,date)};
}

/* ===================== LOOKUP KODE ===================== */
function lookupBookingByCode_(code) {
  const clean=safeString_(code).toUpperCase();
  if(!/^KNR-[A-Z0-9]{5,10}$/.test(clean)) return {success:true,bookings:[]};
  const sheet=getSheet_(); ensureSchema_(sheet); const last=sheet.getLastRow();
  if(last<2) return {success:true,bookings:[]};
  const rows=sheet.getRange(2,1,last-1,HEADERS.length).getValues(), bookings=[];
  rows.forEach(function(r){if(safeString_(r[0]).toUpperCase()===clean) bookings.push(makeBookingObject_(r));});
  bookings.sort(function(a,b){return String(b.createdAt||"").localeCompare(String(a.createdAt||""));});
  return {success:true,bookings:bookings};
}
function makeBookingObject_(r) {
  const created=(Object.prototype.toString.call(r[10])==="[object Date]" && !isNaN(r[10]))?r[10].toISOString():safeString_(r[10]);
  return {
    code:safeString_(r[0]), name:safeString_(r[1]), phone:normalizePhone_(r[2]),
    services:safeString_(r[3]).split(",").map(function(s){return s.trim();}).filter(Boolean),
    barberName:safeString_(r[4]), date:normalizeSheetDate_(r[5]), dateLabel:formatDateLabel_(r[5]),
    time:normalizeSheetTime_(r[6]), total:Number(r[7])||0, notes:safeString_(r[8]),
    status:normalizeStatus_(r[9]), createdAt:created
  };
}

/* ===================== OUTPUT ===================== */
function getErrorMessage_(err) { return err && err.message ? err.message : "Terjadi kesalahan server."; }
function jsonOutput_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function testConnection() {
  const sheet=getSheet_(); ensureSchema_(sheet);
  return {success:true,message:"Koneksi Google Sheets berhasil.",sheet:SHEET_NAME,columns:HEADERS.length};
}
