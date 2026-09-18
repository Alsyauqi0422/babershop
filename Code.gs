const SPREADSHEET_ID = "12-5YwJTdQ4G73QzLM6kVgxAcJ9gZnwzcpU20Xdk1gIs";
const SHEET_NAME = "Reservasi";
const TIMEZONE = "Asia/Makassar";

const HEADERS = [
  "Kode","Nama","WhatsApp","Layanan","Barber",
  "Tanggal","Jam","Total","Catatan","Status","Dibuat"
];

const STATUS_OPTIONS = [
  "Menunggu Konfirmasi",
  "Dikonfirmasi",
  "Selesai",
  "Dibatalkan"
];

const BARBERS = {
  bj: "Bang Jaka",
  kr: "Kang Rudi"
};

const BARBER_ALIASES = {
  b1: "bj",
  b2: "kr",
  bj: "bj",
  kr: "kr"
};

/* =========================
   API
========================= */

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = safeString_(p.action).toLowerCase();

    if (action === "health") {
      return jsonOutput_({
        success: true,
        message: "API Mannuruki Barbershop aktif.",
        endpoints: ["health","slots","lookup"]
      });
    }

    if (action === "slots") {
      return jsonOutput_(getSlots_(p.barberId, p.date));
    }

    if (action === "lookup") {
      return jsonOutput_(lookupBookingByCode_(p.code));
    }

    return jsonOutput_({
      success: true,
      message: "API Mannuruki Barbershop aktif.",
      endpoints: ["health","slots","lookup"]
    });

  } catch (err) {
    return jsonOutput_({
      success: false,
      message: getErrorMessage_(err)
    });
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
      return jsonOutput_({
        success: true,
        duplicate: true,
        message: "Reservasi dengan kode tersebut sudah tersimpan.",
        code: data.code
      });
    }

    const booked = getBookedSlots_(sheet, data.barberId, data.date);

    if (booked.indexOf(data.time) !== -1) {
      return jsonOutput_({
        success: false,
        slotTaken: true,
        message: "Jam tersebut sudah dipesan. Silakan pilih jam lain."
      });
    }

    const services = data.services
      .map(function(s) { return safeString_(s); })
      .filter(Boolean)
      .join(", ");

    const now = new Date();

    const row = [
      data.code,
      data.name,
      data.phone,
      services,
      data.barberName,
      data.date,
      data.time,
      Number(data.total),
      data.notes,
      "Menunggu Konfirmasi",
      now
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    const rowNumber = sheet.getLastRow();
    sheet.getRange(rowNumber, 8).setNumberFormat('"Rp" #,##0');
    sheet.getRange(rowNumber, 11).setNumberFormat("dd/mm/yyyy hh:mm:ss");

    return jsonOutput_({
      success: true,
      message: "Reservasi berhasil disimpan ke Google Sheets.",
      code: data.code,
      booking: makeBookingObject_(row)
    });

  } catch (err) {
    return jsonOutput_({
      success: false,
      message: getErrorMessage_(err)
    });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* =========================
   SHEET
========================= */

function setupSheet() {
  const sheet = getSheet_();

  // PENTING:
  // Fungsi ini membaca struktur yang sedang ada,
  // termasuk struktur 12 kolom dengan Status ganda.
  normalizeSheetStructure_(sheet);
  configureSheet_(sheet);

  return "Sheet berhasil diperbaiki menjadi format A:K.";
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  return sheet;
}

/*
 * MIGRASI ROBUST:
 *
 * Mendukung:
 *
 * FORMAT FINAL:
 * A Kode
 * B Nama
 * C WhatsApp
 * D Layanan
 * E Barber
 * F Tanggal
 * G Jam
 * H Total
 * I Catatan
 * J Status
 * K Dibuat
 *
 * FORMAT LAMA:
 * A Kode
 * B Nama
 * C WhatsApp
 * D Layanan
 * E Barber
 * F Barber ID
 * G Tanggal
 * H Jam
 * I Total
 * J Catatan
 * K Status
 * L Dibuat
 *
 * FORMAT YANG USER KIRIM:
 * A Kode
 * B Nama
 * C WhatsApp
 * D Layanan
 * E Barber
 * F Barber/Barber ID
 * G Tanggal
 * H Jam
 * I Total
 * J Status kosong
 * K Status
 * L Dibuat
 *
 * Fungsi ini mendeteksi berdasarkan isi/header, bukan hanya
 * posisi header, sehingga tidak lagi menghasilkan
 * "Struktur sheet tidak dikenali".
 */
function normalizeSheetStructure_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    return;
  }

  const header = sheet.getRange(1,1,1,Math.max(lastCol,12))
    .getDisplayValues()[0]
    .map(function(v) { return safeString_(v).toLowerCase(); });

  const normalizedHeader = header.map(function(v) {
    return v.replace(/\s+/g, " ").trim();
  });

  // Jika sudah benar A:K, cukup pastikan header.
  if (isFinalHeader_(normalizedHeader)) {
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    return;
  }

  // Jika ada data, baca seluruh data sebelum mengubah kolom.
  const dataRows = lastRow >= 2
    ? sheet.getRange(2,1,lastRow-1,Math.max(lastCol,12)).getValues()
    : [];

  const newRows = dataRows.map(function(r) {
    return convertAnyRowToFinal_(r, normalizedHeader);
  });

  // Pastikan jumlah kolom minimal 12 sebelum membersihkan.
  while (sheet.getMaxColumns() < 12) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  // Bersihkan isi dan format area A:L agar sisa struktur lama hilang.
  sheet.getRange(1,1,sheet.getMaxRows(),12).clearContent();
  sheet.getRange(1,1,sheet.getMaxRows(),12).clearDataValidations();

  // Tulis struktur final.
  sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);

  if (newRows.length) {
    sheet.getRange(2,1,newRows.length,HEADERS.length).setValues(newRows);
  }

  // Hapus kolom setelah K.
  while (sheet.getMaxColumns() > HEADERS.length) {
    sheet.deleteColumn(HEADERS.length + 1);
  }

  SpreadsheetApp.flush();
}

function isFinalHeader_(h) {
  if (h.length < 11) return false;

  const expected = HEADERS.map(function(x) {
    return x.toLowerCase();
  });

  for (let i = 0; i < expected.length; i++) {
    if (safeString_(h[i]) !== expected[i]) return false;
  }

  return true;
}

function convertAnyRowToFinal_(r, h) {
  /*
   * Pertama cari kolom berdasarkan nama header.
   * Ini membuat migrasi aman walaupun posisi kolom lama berbeda.
   */
  const find = function(names) {
    for (let i = 0; i < h.length; i++) {
      if (names.indexOf(h[i]) !== -1) return i;
    }
    return -1;
  };

  let iKode = find(["kode"]);
  let iNama = find(["nama"]);
  let iPhone = find(["whatsapp","phone","nomor whatsapp"]);
  let iLayanan = find(["layanan","service"]);
  let iBarber = find(["barber"]);
  let iTanggal = find(["tanggal","date"]);
  let iJam = find(["jam","time"]);
  let iTotal = find(["total"]);
  let iCatatan = find(["catatan","notes"]);
  let iStatus = find(["status"]);
  let iDibuat = find(["dibuat","created","created at"]);

  /*
   * Jika header tidak membantu, gunakan pola posisi
   * yang diketahui dari sheet user.
   */
  if (iKode < 0) iKode = 0;
  if (iNama < 0) iNama = 1;
  if (iPhone < 0) iPhone = 2;
  if (iLayanan < 0) iLayanan = 3;
  if (iBarber < 0) iBarber = 4;
  if (iTanggal < 0) iTanggal = guessDateColumn_(r);
  if (iJam < 0) iJam = guessTimeColumn_(r);
  if (iTotal < 0) iTotal = guessTotalColumn_(r);

  /*
   * Untuk format 12 kolom lama:
   * F = Barber ID, G = Tanggal, H = Jam, I = Total,
   * J = Catatan, K = Status, L = Dibuat.
   */
  const is12Old =
    r.length >= 12 &&
    isDateLike_(r[6]) &&
    isTimeLike_(r[7]);

  /*
   * Untuk format user:
   * F = barber ID/nama, G = tanggal, H = jam,
   * I = total, J/K = status, L = dibuat.
   */
  const isUser12 =
    r.length >= 12 &&
    isDateLike_(r[6]) &&
    isTimeLike_(r[7]);

  if (is12Old || isUser12) {
    iTanggal = 6;
    iJam = 7;
    iTotal = 8;

    // Pada format lama J adalah Catatan, K adalah Status, L Dibuat.
    // Pada format user J/K sama-sama Status; pilih nilai status yang valid.
    const possibleJ = safeString_(r[9]);
    const possibleK = safeString_(r[10]);

    if (isStatus_(possibleK)) {
      iStatus = 10;
      iCatatan = 9;
    } else if (isStatus_(possibleJ)) {
      iStatus = 9;
      iCatatan = -1;
    } else {
      iStatus = -1;
      iCatatan = 9;
    }

    iDibuat = 11;
  }

  const status =
    iStatus >= 0 && isStatus_(r[iStatus])
      ? normalizeStatus_(r[iStatus])
      : "Menunggu Konfirmasi";

  const barberValue = safeString_(r[iBarber]);

  const barberName = resolveBarberName_(barberValue);

  const catatan =
    iCatatan >= 0
      ? safeString_(r[iCatatan]).slice(0,500)
      : "";

  const dibuat =
    iDibuat >= 0
      ? normalizeCreatedAt_(r[iDibuat])
      : new Date();

  return [
    safeString_(r[iKode]).toUpperCase(),
    safeString_(r[iNama]),
    normalizePhone_(r[iPhone]),
    safeString_(r[iLayanan]),
    barberName,
    normalizeSheetDate_(r[iTanggal]),
    normalizeSheetTime_(r[iJam]),
    normalizeNumber_(r[iTotal]),
    catatan,
    status,
    dibuat
  ];
}

function resolveBarberName_(value) {
  const s = safeString_(value);

  if (!s) return "";

  const lower = s.toLowerCase();

  if (BARBER_ALIASES[lower]) {
    return BARBERS[BARBER_ALIASES[lower]];
  }

  if (lower === "bang jaka") return "Bang Jaka";
  if (lower === "kang rudi") return "Kang Rudi";

  return s;
}

function guessDateColumn_(r) {
  for (let i = 0; i < r.length; i++) {
    if (isDateLike_(r[i])) return i;
  }
  return 5;
}

function guessTimeColumn_(r) {
  for (let i = 0; i < r.length; i++) {
    if (isTimeLike_(r[i])) return i;
  }
  return 6;
}

function guessTotalColumn_(r) {
  // Cari angka yang masuk akal sebagai total, dari kanan ke kiri.
  for (let i = r.length - 1; i >= 0; i--) {
    const n = Number(r[i]);
    if (Number.isFinite(n) && n > 0 && n <= 100000000) {
      return i;
    }
  }
  return 7;
}

function configureSheet_(sheet) {
  sheet.getRange(1,1,1,HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("F:G").setNumberFormat("@");
  sheet.getRange("H:H").setNumberFormat('"Rp" #,##0');
  sheet.getRange("K:K").setNumberFormat("dd/mm/yyyy hh:mm:ss");

  // Hapus semua validasi lama terlebih dahulu.
  sheet.getRange("A:K").clearDataValidations();

  // Validasi STATUS hanya J2:J.
  const statusRange = sheet.getRange(
    2,
    10,
    Math.max(sheet.getMaxRows() - 1, 1),
    1
  );

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .setHelpText(
      "Pilih: Menunggu Konfirmasi, Dikonfirmasi, Selesai, atau Dibatalkan."
    )
    .build();

  statusRange.setDataValidation(rule);

  const widths = [130,180,140,300,140,130,90,130,250,180,170];

  widths.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
}

/* =========================
   VALIDASI BOOKING
========================= */

function parsePostData_(e) {
  if (!e) throw new Error("Data reservasi tidak ditemukan.");

  let raw = "";

  if (e.parameter && e.parameter.data) {
    raw = e.parameter.data;
  } else if (e.postData && e.postData.contents) {
    raw = e.postData.contents;
  }

  if (!raw) throw new Error("Data reservasi tidak ditemukan.");

  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") throw new Error();
    return parsed;
  } catch (_) {
    throw new Error("Format data reservasi tidak valid.");
  }
}

function validateBooking_(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Format data reservasi tidak valid.");
  }

  data.code = safeString_(data.code).toUpperCase();

  if (!/^KNR-[A-Z0-9]{5,10}$/.test(data.code)) {
    throw new Error("Kode reservasi tidak valid.");
  }

  data.name = safeString_(data.name);

  if (data.name.length < 3 || data.name.length > 100) {
    throw new Error("Nama pelanggan harus 3–100 karakter.");
  }

  data.phone = normalizePhone_(data.phone);

  if (data.phone.length < 9 || data.phone.length > 15) {
    throw new Error("Nomor WhatsApp tidak valid.");
  }

  data.barberId = normalizeBarberId_(data.barberId);

  if (!BARBERS[data.barberId]) {
    const barberNameInput = safeString_(data.barberName).toLowerCase();

    if (barberNameInput === "bang jaka") data.barberId = "bj";
    if (barberNameInput === "kang rudi") data.barberId = "kr";
  }

  if (!BARBERS[data.barberId]) {
    throw new Error("Barber belum dipilih.");
  }

  data.barberName = BARBERS[data.barberId];

  if (!isValidDate_(data.date)) {
    throw new Error("Tanggal reservasi tidak valid.");
  }

  if (!isValidTime_(data.time)) {
    throw new Error("Jam reservasi tidak valid.");
  }

  data.services = Array.isArray(data.services)
    ? data.services.map(function(s) {
        return safeString_(s);
      }).filter(Boolean).slice(0,10)
    : [];

  if (!data.services.length) {
    throw new Error("Layanan belum dipilih.");
  }

  data.notes = safeString_(data.notes).slice(0,500);
  data.total = Number(data.total);

  if (
    !Number.isFinite(data.total) ||
    data.total < 0 ||
    data.total > 100000000
  ) {
    throw new Error("Total reservasi tidak valid.");
  }
}

/* =========================
   NORMALISASI
========================= */

function normalizePhone_(value) {
  let d = safeString_(value).replace(/\D/g, "");

  if (!d) return "";

  if (d.indexOf("0") === 0) {
    return "62" + d.slice(1);
  }

  if (d.indexOf("62") === 0) {
    return d;
  }

  return d;
}

function normalizeBarberId_(id) {
  const s = safeString_(id).toLowerCase();
  return BARBER_ALIASES[s] || s;
}

function safeString_(value) {
  return value == null ? "" : String(value).trim();
}

function normalizeNumber_(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const s = safeString_(value).replace(/[^\d.-]/g, "");
  const n = Number(s);

  return Number.isFinite(n) ? n : 0;
}

function normalizeCreatedAt_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) {
    return value;
  }

  const s = safeString_(value);

  if (!s) return new Date();

  const d = new Date(s);

  return isNaN(d) ? new Date() : d;
}

function isStatus_(value) {
  const s = safeString_(value).toLowerCase();

  return [
    "menunggu konfirmasi",
    "menunggu",
    "dikonfirmasi",
    "selesai",
    "dibatalkan",
    "cancelled",
    "canceled"
  ].indexOf(s) !== -1;
}

function normalizeStatus_(value) {
  const s = safeString_(value).toLowerCase();

  if (s === "dikonfirmasi") return "Dikonfirmasi";
  if (s === "selesai") return "Selesai";
  if (
    s === "dibatalkan" ||
    s === "cancelled" ||
    s === "canceled"
  ) return "Dibatalkan";

  return "Menunggu Konfirmasi";
}

function isCancelledStatus_(status) {
  return normalizeStatus_(status) === "Dibatalkan";
}

/* =========================
   DATE/TIME
========================= */

function isValidDate_(value) {
  const s = safeString_(value);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;

  const p = s.split("-").map(Number);
  const dt = new Date(p[0], p[1] - 1, p[2]);

  return (
    dt.getFullYear() === p[0] &&
    dt.getMonth() === p[1] - 1 &&
    dt.getDate() === p[2]
  );
}

function isValidTime_(value) {
  const s = safeString_(value);

  if (!/^\d{2}:\d{2}$/.test(s)) return false;

  const p = s.split(":").map(Number);

  return p[0] >= 0 && p[0] <= 23 && p[1] >= 0 && p[1] <= 59;
}

function isDateLike_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) return true;

  const s = safeString_(value);

  return /^\d{4}-\d{2}-\d{2}$/.test(s) ||
         /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s);
}

function isTimeLike_(value) {
  const s = safeString_(value);
  return /^\d{1,2}:\d{2}/.test(s);
}

function normalizeSheetDate_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(value, TIMEZONE, "yyyy-MM-dd");
  }

  const s = safeString_(value);

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);

  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);

    const result =
      y + "-" +
      String(mo).padStart(2,"0") + "-" +
      String(d).padStart(2,"0");

    if (isValidDate_(result)) return result;
  }

  const d = new Date(s);

  return isNaN(d)
    ? s
    : Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd");
}

function normalizeSheetTime_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(value, TIMEZONE, "HH:mm");
  }

  const s = safeString_(value);
  const m = s.match(/^(\d{1,2}):(\d{2})/);

  if (!m) return s;

  return String(Number(m[1])).padStart(2,"0") + ":" + m[2];
}

/* =========================
   SLOTS
========================= */

function getSlots_(barberId, date) {
  barberId = normalizeBarberId_(barberId);
  date = safeString_(date);

  if (!BARBERS[barberId] || !isValidDate_(date)) {
    return { success:true, slots:[] };
  }

  const sheet = getSheet_();
  ensureSchema_(sheet);

  return {
    success:true,
    slots:getBookedSlots_(sheet, barberId, date)
  };
}

function getBookedSlots_(sheet, barberId, date) {
  const last = sheet.getLastRow();

  if (last < 2) return [];

  const barberName =
    BARBERS[normalizeBarberId_(barberId)] || "";

  const rows = sheet
    .getRange(2,1,last-1,HEADERS.length)
    .getValues();

  const out = [];

  rows.forEach(function(r) {
    const storedBarber = safeString_(r[4]);
    const storedDate = normalizeSheetDate_(r[5]);
    const storedTime = normalizeSheetTime_(r[6]);
    const status = safeString_(r[9]);

    if (
      storedBarber === barberName &&
      storedDate === date &&
      storedTime &&
      !isCancelledStatus_(status)
    ) {
      out.push(storedTime);
    }
  });

  return Array.from(new Set(out)).sort();
}

/* =========================
   LOOKUP
========================= */

function lookupBookingByCode_(code) {
  const clean = safeString_(code).toUpperCase();

  if (!/^KNR-[A-Z0-9]{5,10}$/.test(clean)) {
    return { success:true, bookings:[] };
  }

  const sheet = getSheet_();
  ensureSchema_(sheet);

  const last = sheet.getLastRow();

  if (last < 2) {
    return { success:true, bookings:[] };
  }

  const rows = sheet
    .getRange(2,1,last-1,HEADERS.length)
    .getValues();

  const bookings = [];

  rows.forEach(function(r) {
    if (safeString_(r[0]).toUpperCase() === clean) {
      bookings.push(makeBookingObject_(r));
    }
  });

  bookings.sort(function(a,b) {
    return String(b.createdAt || "")
      .localeCompare(String(a.createdAt || ""));
  });

  return {
    success:true,
    bookings:bookings
  };
}

function makeBookingObject_(r) {
  const created =
    r[10] instanceof Date
      ? r[10].toISOString()
      : safeString_(r[10]);

  return {
    code:safeString_(r[0]),
    name:safeString_(r[1]),
    phone:normalizePhone_(r[2]),
    services:safeString_(r[3])
      .split(",")
      .map(function(s) { return s.trim(); })
      .filter(Boolean),
    barberName:safeString_(r[4]),
    date:normalizeSheetDate_(r[5]),
    dateLabel:formatDateLabel_(r[5]),
    time:normalizeSheetTime_(r[6]),
    total:Number(r[7]) || 0,
    notes:safeString_(r[8]),
    status:normalizeStatus_(r[9]),
    createdAt:created
  };
}

function formatDateLabel_(value) {
  const normalized = normalizeSheetDate_(value);

  if (!isValidDate_(normalized)) {
    return safeString_(value);
  }

  const p = normalized.split("-").map(Number);

  const names = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];

  return p[2] + " " + names[p[1]-1] + " " + p[0];
}

/* =========================
   UTIL
========================= */

function codeExists_(sheet, code) {
  const last = sheet.getLastRow();

  if (last < 2) return false;

  const target = safeString_(code).toUpperCase();

  return sheet
    .getRange(2,1,last-1,1)
    .getValues()
    .some(function(r) {
      return safeString_(r[0]).toUpperCase() === target;
    });
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(err) {
  return err && err.message
    ? err.message
    : "Terjadi kesalahan server.";
}

function testConnection() {
  const sheet = getSheet_();
  normalizeSheetStructure_(sheet);
  configureSheet_(sheet);

  return {
    success:true,
    message:"Koneksi Google Sheets berhasil.",
    sheet:SHEET_NAME,
    columns:HEADERS.length
  };
}
