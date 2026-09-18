const SPREADSHEET_ID = "12-5YwJTdQ4G73QzLM6kVgxAcJ9gZnwzcpU20Xdk1gIs";
const SHEET_NAME = "Reservasi";
const TIMEZONE = "Asia/Makassar";
const HEADERS = ["Kode","Nama","WhatsApp","Layanan","Barber","Tanggal","Jam","Total","Catatan","Status","Dibuat"];
const BARBERS = { bj: "Bang Jaka", kr: "Kang Rudi" };

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = safeString_(p.action).toLowerCase();
    if (action === "slots") return jsonOutput_(getSlots_(p.barberId, p.date));
    if (action === "lookup") return jsonOutput_(lookupBookingByCode_(p.code));
    return jsonOutput_({ success:true, message:"API Mannuruki Barbershop aktif.", endpoints:["slots","lookup"] });
  } catch (err) {
    return jsonOutput_({ success:false, message:err.message || "Terjadi kesalahan server." });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const data = parsePostData_(e);
    validateBooking_(data);
    const sheet = getSheet_();
    ensureHeader_(sheet);

    if (codeExists_(sheet, data.code)) {
      return jsonOutput_({success:true, duplicate:true, message:"Reservasi dengan kode tersebut sudah tersimpan.", code:data.code});
    }

    const booked = getBookedSlots_(sheet, data.barberId, data.date);
    if (booked.indexOf(data.time) !== -1) {
      return jsonOutput_({success:false, slotTaken:true, message:"Jam tersebut sudah dipesan. Silakan pilih jam lain."});
    }

    const services = data.services.map(s => safeString_(s)).filter(Boolean).join(", ");
    const total = Number(data.total);
    const now = new Date();
    const row = [
      data.code, data.name, normalizePhone_(data.phone), services,
      data.barberName, data.date, data.time, total, data.notes,
      "Menunggu Konfirmasi", now
    ];
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    const rowNumber = sheet.getLastRow();
    sheet.getRange(rowNumber, 8).setNumberFormat('"Rp" #,##0');
    sheet.getRange(rowNumber, 11).setNumberFormat("dd/mm/yyyy hh:mm:ss");

    const booking = makeBookingObject_(row);
    return jsonOutput_({success:true, message:"Reservasi berhasil disimpan ke Google Sheets.", code:data.code, booking:booking});
  } catch (err) {
    return jsonOutput_({success:false, message:err.message || "Reservasi gagal disimpan."});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function setupSheet() {
  const sheet = getSheet_();
  ensureHeader_(sheet);
  sheet.setFrozenRows(1);
  sheet.getRange("H:H").setNumberFormat('"Rp" #,##0');
  sheet.getRange("K:K").setNumberFormat("dd/mm/yyyy hh:mm:ss");
  const widths = [130,180,140,300,140,130,90,130,250,180,170];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1,w));
  return "Sheet berhasil disiapkan.";
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeader_(sheet) {
  const current = sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  let different = false;
  for (let i=0;i<HEADERS.length;i++) if (String(current[i] || "").trim() !== HEADERS[i]) { different=true; break; }
  if (different) sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1,1,1,HEADERS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("F:G").setNumberFormat("@");
}

function parsePostData_(e) {
  if (!e) throw new Error("Data reservasi tidak ditemukan.");
  let raw = "";
  if (e.parameter && e.parameter.data) raw = e.parameter.data;
  else if (e.postData && e.postData.contents) raw = e.postData.contents;
  if (!raw) throw new Error("Data reservasi tidak ditemukan.");
  try { return JSON.parse(raw); } catch (_) { throw new Error("Format data reservasi tidak valid."); }
}

function validateBooking_(data) {
  if (!data || typeof data !== "object") throw new Error("Format data reservasi tidak valid.");
  const code = safeString_(data.code).toUpperCase();
  if (!/^KNR-[A-Z0-9]{5,10}$/.test(code)) throw new Error("Kode reservasi tidak valid.");
  data.code = code;
  data.name = safeString_(data.name);
  if (data.name.length < 3 || data.name.length > 100) throw new Error("Nama pelanggan harus 3–100 karakter.");
  data.phone = normalizePhone_(data.phone);
  if (data.phone.length < 9 || data.phone.length > 15) throw new Error("Nomor WhatsApp tidak valid.");
  data.barberId = normalizeBarberId_(data.barberId);
  if (!BARBERS[data.barberId]) throw new Error("Barber belum dipilih.");
  data.barberName = BARBERS[data.barberId];
  if (!isValidDate_(data.date)) throw new Error("Tanggal reservasi tidak valid.");
  if (!isValidTime_(data.time)) throw new Error("Jam reservasi tidak valid.");
  data.services = Array.isArray(data.services) ? data.services.map(s=>safeString_(s)).filter(Boolean).slice(0,10) : [];
  if (!data.services.length) throw new Error("Layanan belum dipilih.");
  data.notes = safeString_(data.notes).slice(0,500);
  data.total = Number(data.total);
  if (!Number.isFinite(data.total) || data.total < 0 || data.total > 100000000) throw new Error("Total reservasi tidak valid.");
}

function normalizePhone_(value) {
  let d = safeString_(value).replace(/\D/g, "");
  if (d.indexOf("+62") === 0) d = d.slice(1);
  if (d.indexOf("62") === 0) return "62" + d.slice(2);
  if (d.indexOf("0") === 0) return "62" + d.slice(1);
  return d;
}

function normalizeBarberId_(id) { return safeString_(id).toLowerCase(); }
function safeString_(value) { return value == null ? "" : String(value).trim(); }
function isCancelledStatus_(status) { const s=safeString_(status).toLowerCase(); return s === "dibatalkan" || s === "cancelled" || s === "canceled"; }

function isValidDate_(value) {
  const s = safeString_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const [y,m,d] = s.split("-").map(Number);
  const dt = new Date(y,m-1,d);
  return dt.getFullYear()===y && dt.getMonth()===m-1 && dt.getDate()===d;
}

function isValidTime_(value) {
  const s=safeString_(value);
  if (!/^\d{2}:\d{2}$/.test(s)) return false;
  const [h,m]=s.split(":").map(Number);
  return h>=0 && h<=23 && m>=0 && m<=59;
}

function codeExists_(sheet, code) {
  const last=sheet.getLastRow(); if(last<2) return false;
  const target=safeString_(code).toUpperCase();
  return sheet.getRange(2,1,last-1,1).getValues().some(r=>safeString_(r[0]).toUpperCase()===target);
}

function getBookedSlots_(sheet, barberId, date) {
  const last=sheet.getLastRow(); if(last<2) return [];
  const barberName=BARBERS[normalizeBarberId_(barberId)] || "";
  const rows=sheet.getRange(2,1,last-1,11).getValues();
  const out=[];
  rows.forEach(r=>{
    const storedBarber=safeString_(r[4]);
    const storedDate=normalizeSheetDate_(r[5]);
    const storedTime=normalizeSheetTime_(r[6]);
    if(storedBarber===barberName && storedDate===date && storedTime && !isCancelledStatus_(r[9])) out.push(storedTime);
  });
  return [...new Set(out)].sort();
}

function normalizeSheetDate_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return Utilities.formatDate(value,TIMEZONE,"yyyy-MM-dd");
  const s=safeString_(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d=new Date(s); return isNaN(d) ? s : Utilities.formatDate(d,TIMEZONE,"yyyy-MM-dd");
}
function normalizeSheetTime_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return Utilities.formatDate(value,TIMEZONE,"HH:mm");
  const s=safeString_(value);
  const m=s.match(/^(\d{1,2}):(\d{2})/);
  return m ? String(Number(m[1])).padStart(2,"0")+":"+m[2] : s;
}

function getSlots_(barberId,date) {
  barberId=normalizeBarberId_(barberId); date=safeString_(date);
  if(!BARBERS[barberId] || !isValidDate_(date)) return {success:true,slots:[]};
  const sheet=getSheet_(); ensureHeader_(sheet);
  return {success:true,slots:getBookedSlots_(sheet,barberId,date)};
}

function lookupBookingByCode_(code) {
  const clean=safeString_(code).toUpperCase();
  if(!/^KNR-[A-Z0-9]{5,10}$/.test(clean)) return {success:true,bookings:[]};
  const sheet=getSheet_(); ensureHeader_(sheet);
  const last=sheet.getLastRow(); if(last<2) return {success:true,bookings:[]};
  const rows=sheet.getRange(2,1,last-1,11).getValues();
  const bookings=[];
  rows.forEach(r=>{ if(safeString_(r[0]).toUpperCase()===clean) bookings.push(makeBookingObject_(r)); });
  bookings.sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));
  return {success:true,bookings:bookings};
}

function makeBookingObject_(r) {
  const created = r[10] instanceof Date ? r[10].toISOString() : safeString_(r[10]);
  return {
    code:safeString_(r[0]), name:safeString_(r[1]), phone:normalizePhone_(r[2]),
    services:safeString_(r[3]).split(",").map(s=>s.trim()).filter(Boolean),
    barberName:safeString_(r[4]), date:r[5] instanceof Date ? Utilities.formatDate(r[5],TIMEZONE,"yyyy-MM-dd") : safeString_(r[5]),
    dateLabel:formatDateLabel_(r[5]), time:normalizeSheetTime_(r[6]), total:Number(r[7])||0,
    notes:safeString_(r[8]), status:safeString_(r[9]) || "Menunggu Konfirmasi", createdAt:created
  };
}

function formatDateLabel_(value) {
  if (Object.prototype.toString.call(value)==="[object Date]" && !isNaN(value)) return Utilities.formatDate(value,TIMEZONE,"dd MMMM yyyy");
  const s=safeString_(value); if(isValidDate_(s)){ const [y,m,d]=s.split("-").map(Number); const names=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"]; return d+" "+names[m-1]+" "+y; }
  return s;
}

function jsonOutput_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
