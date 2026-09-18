const SPREADSHEET_ID = "12-5YwJTdQ4G73QzLM6kVgxAcJ9gZnwzcpU20Xdk1gIs";
const SHEET_NAME = "Reservasi";
const TIMEZONE = "Asia/Makassar";

/* =========================
   KONFIGURASI SHEET
========================= */

const HEADERS = [
  "Kode",
  "Nama",
  "WhatsApp",
  "Layanan",
  "Barber",
  "Tanggal",
  "Jam",
  "Total",
  "Catatan",
  "Status",
  "Dibuat"
];

const STATUS_OPTIONS = [
  "Menunggu Konfirmasi",
  "Dikonfirmasi",
  "Selesai",
  "Dibatalkan"
];

/* =========================
   DATA BARBER
========================= */

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
   API GET
========================= */

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = safeString_(p.action).toLowerCase();

    /* ---------- HEALTH ---------- */

    if (action === "health") {
      return jsonOutput_({
        success: true,
        message: "API Mannuruki Barbershop aktif.",
        endpoints: [
          "health",
          "slots",
          "lookup"
        ]
      });
    }

    /* ---------- SLOTS ---------- */

    if (action === "slots") {
      return jsonOutput_(
        getSlots_(p.barberId, p.date)
      );
    }

    /* ---------- LOOKUP ---------- */

    if (action === "lookup") {
      return jsonOutput_(
        lookupBookingByCode_(p.code)
      );
    }

    /* ---------- DEFAULT ---------- */

    return jsonOutput_({
      success: true,
      message: "API Mannuruki Barbershop aktif.",
      endpoints: [
        "health",
        "slots",
        "lookup"
      ]
    });

  } catch (err) {

    return jsonOutput_({
      success: false,
      message: getErrorMessage_(err)
    });

  }
}

/* =========================
   API POST
========================= */

function doPost(e) {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(15000);

    /* ---------- BACA DATA ---------- */

    const data = parsePostData_(e);

    /* ---------- VALIDASI ---------- */

    validateBooking_(data);

    /* ---------- AMBIL SHEET ---------- */

    const sheet = getSheet_();

    /*
     * TIDAK menggunakan ensureSchema_().
     *
     * Struktur sheet sudah disiapkan melalui
     * fungsi setupSheet().
     */

    /* ---------- CEK KODE DUPLIKAT ---------- */

    if (codeExists_(sheet, data.code)) {

      return jsonOutput_({
        success: true,
        duplicate: true,
        message: "Reservasi dengan kode tersebut sudah tersimpan.",
        code: data.code
      });

    }

    /* ---------- CEK JAM SUDAH DIPESAN ---------- */

    const booked = getBookedSlots_(
      sheet,
      data.barberId,
      data.date
    );

    if (booked.indexOf(data.time) !== -1) {

      return jsonOutput_({
        success: false,
        slotTaken: true,
        message: "Jam tersebut sudah dipesan. Silakan pilih jam lain."
      });

    }

    /* ---------- FORMAT LAYANAN ---------- */

    const services = data.services
      .map(function(service) {
        return safeString_(service);
      })
      .filter(Boolean)
      .join(", ");

    /* ---------- WAKTU DIBUAT ---------- */

    const now = new Date();

    /* ---------- DATA ROW ---------- */

    const row = [
      data.code,             // A Kode
      data.name,             // B Nama
      data.phone,            // C WhatsApp
      services,              // D Layanan
      data.barberName,       // E Barber
      data.date,             // F Tanggal
      data.time,             // G Jam
      Number(data.total),    // H Total
      data.notes,            // I Catatan
      "Menunggu Konfirmasi", // J Status
      now                    // K Dibuat
    ];

    /* ---------- SIMPAN ---------- */

    sheet.appendRow(row);

    SpreadsheetApp.flush();

    /* ---------- FORMAT BARIS ---------- */

    const rowNumber = sheet.getLastRow();

    sheet
      .getRange(rowNumber, 8)
      .setNumberFormat('"Rp" #,##0');

    sheet
      .getRange(rowNumber, 11)
      .setNumberFormat("dd/mm/yyyy hh:mm:ss");

    /* ---------- RESPONSE ---------- */

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

    try {
      lock.releaseLock();
    } catch (_) {}

  }
}

/* =========================
   SETUP SHEET
========================= */

/*
 * JALANKAN FUNGSI INI SEKALI SAJA
 * SETELAH CODE.GS DITEMPEL.
 *
 * Fungsi ini akan:
 * - memperbaiki header
 * - menghapus validasi lama
 * - memperbaiki struktur 12 kolom
 * - mengubah menjadi A:K
 * - mempertahankan data reservasi
 * - membuat validasi Status di J
 * - menghapus kolom setelah K
 */

function setupSheet() {

  const sheet = getSheet_();

  normalizeSheetStructure_(sheet);

  configureSheet_(sheet);

  SpreadsheetApp.flush();

  return "Sheet berhasil diperbaiki menjadi format A:K.";
}

/* =========================
   AMBIL SHEET
========================= */

function getSheet_() {

  const ss = SpreadsheetApp.openById(
    SPREADSHEET_ID
  );

  let sheet = ss.getSheetByName(
    SHEET_NAME
  );

  if (!sheet) {

    sheet = ss.insertSheet(
      SHEET_NAME
    );

  }

  return sheet;
}

/* =========================
   NORMALISASI STRUKTUR
========================= */

function normalizeSheetStructure_(sheet) {

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  /*
   * PENTING:
   * Pastikan minimal tersedia 12 kolom
   * sebelum proses migrasi.
   */

  while (sheet.getMaxColumns() < 12) {

    sheet.insertColumnAfter(
      sheet.getMaxColumns()
    );

  }

  /*
   * HAPUS VALIDASI LAMA TERLEBIH DAHULU.
   *
   * Ini penting untuk menghindari error:
   *
   * "Data yang dimasukkan ke dalam sel K1
   * melanggar aturan validasi..."
   */

  sheet
    .getRange(
      1,
      1,
      sheet.getMaxRows(),
      Math.max(sheet.getMaxColumns(), 12)
    )
    .clearDataValidations();

  /* ---------- SHEET KOSONG ---------- */

  if (lastRow === 0 || lastCol === 0) {

    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);

    return;

  }

  /*
   * BACA HEADER LAMA
   */

  const readCols = Math.max(
    lastCol,
    12
  );

  const header = sheet
    .getRange(
      1,
      1,
      1,
      readCols
    )
    .getDisplayValues()[0]
    .map(function(value) {

      return safeString_(value)
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();

    });

  /*
   * JIKA SUDAH FORMAT FINAL A:K
   */

  if (isFinalHeader_(header)) {

    sheet
      .getRange(
        1,
        1,
        1,
        HEADERS.length
      )
      .setValues([HEADERS]);

    return;

  }

  /*
   * BACA DATA LAMA
   */

  const dataRows = lastRow >= 2

    ? sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          readCols
        )
        .getValues()

    : [];

  /*
   * KONVERSI DATA
   */

  const newRows = dataRows.map(
    function(row) {

      return convertAnyRowToFinal_(
        row,
        header
      );

    }
  );

  /*
   * BERSIHKAN AREA A:L
   */

  sheet
    .getRange(
      1,
      1,
      sheet.getMaxRows(),
      12
    )
    .clearContent();

  sheet
    .getRange(
      1,
      1,
      sheet.getMaxRows(),
      12
    )
    .clearDataValidations();

  /*
   * TULIS HEADER FINAL
   */

  sheet
    .getRange(
      1,
      1,
      1,
      HEADERS.length
    )
    .setValues([HEADERS]);

  /*
   * TULIS DATA
   */

  if (newRows.length) {

    sheet
      .getRange(
        2,
        1,
        newRows.length,
        HEADERS.length
      )
      .setValues(newRows);

  }

  /*
   * HAPUS KOLOM SETELAH K
   */

  while (
    sheet.getMaxColumns() > HEADERS.length
  ) {

    sheet.deleteColumn(
      HEADERS.length + 1
    );

  }

  SpreadsheetApp.flush();
}

/* =========================
   CEK HEADER FINAL
========================= */

function isFinalHeader_(header) {

  if (header.length < 11) {
    return false;
  }

  const expected = HEADERS.map(
    function(value) {
      return value.toLowerCase();
    }
  );

  for (
    let i = 0;
    i < expected.length;
    i++
  ) {

    if (
      safeString_(header[i]) !==
      expected[i]
    ) {

      return false;

    }

  }

  return true;
}

/* =========================
   KONVERSI DATA LAMA
========================= */

function convertAnyRowToFinal_(r, h) {

  /*
   * CARI KOLOM BERDASARKAN HEADER
   */

  const find = function(names) {

    for (
      let i = 0;
      i < h.length;
      i++
    ) {

      if (
        names.indexOf(h[i]) !== -1
      ) {

        return i;

      }

    }

    return -1;
  };

  let iKode = find([
    "kode",
    "code"
  ]);

  let iNama = find([
    "nama",
    "name"
  ]);

  let iPhone = find([
    "whatsapp",
    "phone",
    "nomor whatsapp",
    "nomor",
    "no whatsapp"
  ]);

  let iLayanan = find([
    "layanan",
    "service"
  ]);

  let iBarber = find([
    "barber",
    "capster"
  ]);

  let iTanggal = find([
    "tanggal",
    "date"
  ]);

  let iJam = find([
    "jam",
    "time"
  ]);

  let iTotal = find([
    "total"
  ]);

  let iCatatan = find([
    "catatan",
    "notes"
  ]);

  let iStatus = find([
    "status"
  ]);

  let iDibuat = find([
    "dibuat",
    "created",
    "created at"
  ]);

  /*
   * DEFAULT POSISI
   */

  if (iKode < 0) {
    iKode = 0;
  }

  if (iNama < 0) {
    iNama = 1;
  }

  if (iPhone < 0) {
    iPhone = 2;
  }

  if (iLayanan < 0) {
    iLayanan = 3;
  }

  if (iBarber < 0) {
    iBarber = 4;
  }

  if (iTanggal < 0) {
    iTanggal = guessDateColumn_(r);
  }

  if (iJam < 0) {
    iJam = guessTimeColumn_(r);
  }

  if (iTotal < 0) {
    iTotal = guessTotalColumn_(r);
  }

  /*
   * DETEKSI FORMAT 12 KOLOM LAMA
   *
   * F = Barber ID
   * G = Tanggal
   * H = Jam
   * I = Total
   * J = Catatan / Status
   * K = Status
   * L = Dibuat
   */

  const isOld12 =
    r.length >= 12 &&
    isDateLike_(r[6]) &&
    isTimeLike_(r[7]);

  if (isOld12) {

    /*
     * Pastikan barber menggunakan
     * kolom E terlebih dahulu.
     */

    let barberValue =
      safeString_(r[4]);

    /*
     * Jika E kosong, coba F.
     */

    if (!barberValue) {

      barberValue =
        safeString_(r[5]);

    }

    /*
     * Tanggal, jam, total
     */

    iTanggal = 6;
    iJam = 7;
    iTotal = 8;

    /*
     * J/K kemungkinan status ganda.
     */

    const possibleJ =
      safeString_(r[9]);

    const possibleK =
      safeString_(r[10]);

    /*
     * Tentukan status.
     */

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

    /*
     * Dibuat selalu L
     */

    iDibuat = 11;

    /*
     * Barber value disimpan kembali
     */

    const barberName =
      resolveBarberName_(barberValue);

    return [
      safeString_(r[iKode]).toUpperCase(),
      safeString_(r[iNama]),
      normalizePhone_(r[iPhone]),
      safeString_(r[iLayanan]),
      barberName,
      normalizeSheetDate_(r[iTanggal]),
      normalizeSheetTime_(r[iJam]),
      normalizeNumber_(r[iTotal]),
      iCatatan >= 0
        ? safeString_(r[iCatatan]).slice(0, 500)
        : "",
      iStatus >= 0 &&
      isStatus_(r[iStatus])
        ? normalizeStatus_(r[iStatus])
        : "Menunggu Konfirmasi",
      normalizeCreatedAt_(r[iDibuat])
    ];

  }

  /*
   * FORMAT LAIN / FORMAT UMUM
   */

  const barberValue =
    safeString_(r[iBarber]);

  /*
   * Jika kolom E kosong, coba F
   * karena format lama memiliki
   * Barber ID di F.
   */

  let finalBarberValue =
    barberValue;

  if (
    !finalBarberValue &&
    r.length >= 6
  ) {

    finalBarberValue =
      safeString_(r[5]);

  }

  const barberName =
    resolveBarberName_(
      finalBarberValue
    );

  const status =
    iStatus >= 0 &&
    isStatus_(r[iStatus])

      ? normalizeStatus_(r[iStatus])

      : "Menunggu Konfirmasi";

  const catatan =
    iCatatan >= 0

      ? safeString_(
          r[iCatatan]
        ).slice(0, 500)

      : "";

  const dibuat =
    iDibuat >= 0

      ? normalizeCreatedAt_(
          r[iDibuat]
        )

      : new Date();

  return [

    safeString_(
      r[iKode]
    ).toUpperCase(),

    safeString_(
      r[iNama]
    ),

    normalizePhone_(
      r[iPhone]
    ),

    safeString_(
      r[iLayanan]
    ),

    barberName,

    normalizeSheetDate_(
      r[iTanggal]
    ),

    normalizeSheetTime_(
      r[iJam]
    ),

    normalizeNumber_(
      r[iTotal]
    ),

    catatan,

    status,

    dibuat

  ];
}

/* =========================
   RESOLVE BARBER
========================= */

function resolveBarberName_(value) {

  const s =
    safeString_(value);

  if (!s) {
    return "";
  }

  const lower =
    s.toLowerCase();

  if (
    BARBER_ALIASES[lower]
  ) {

    return BARBERS[
      BARBER_ALIASES[lower]
    ];

  }

  if (
    lower === "bang jaka"
  ) {

    return "Bang Jaka";

  }

  if (
    lower === "kang rudi"
  ) {

    return "Kang Rudi";

  }

  return s;
}

/* =========================
   TEBAK KOLOM TANGGAL
========================= */

function guessDateColumn_(r) {

  for (
    let i = 0;
    i < r.length;
    i++
  ) {

    if (
      isDateLike_(r[i])
    ) {

      return i;

    }

  }

  return 5;
}

/* =========================
   TEBAK KOLOM JAM
========================= */

function guessTimeColumn_(r) {

  for (
    let i = 0;
    i < r.length;
    i++
  ) {

    if (
      isTimeLike_(r[i])
    ) {

      return i;

    }

  }

  return 6;
}

/* =========================
   TEBAK KOLOM TOTAL
========================= */

function guessTotalColumn_(r) {

  /*
   * Cari angka dari kanan ke kiri.
   */

  for (
    let i = r.length - 1;
    i >= 0;
    i--
  ) {

    const n =
      Number(r[i]);

    if (
      Number.isFinite(n) &&
      n > 0 &&
      n <= 100000000
    ) {

      return i;

    }

  }

  return 7;
}

/* =========================
   KONFIGURASI SHEET
========================= */

function configureSheet_(sheet) {

  /*
   * HEADER
   */

  sheet
    .getRange(
      1,
      1,
      1,
      HEADERS.length
    )
    .setValues([HEADERS])
    .setFontWeight("bold");

  /*
   * FREEZE HEADER
   */

  sheet.setFrozenRows(1);

  /*
   * FORMAT KOLOM
   */

  sheet
    .getRange("A:A")
    .setNumberFormat("@");

  sheet
    .getRange("C:C")
    .setNumberFormat("@");

  sheet
    .getRange("F:G")
    .setNumberFormat("@");

  sheet
    .getRange("H:H")
    .setNumberFormat('"Rp" #,##0');

  sheet
    .getRange("K:K")
    .setNumberFormat(
      "dd/mm/yyyy hh:mm:ss"
    );

  /*
   * HAPUS SEMUA VALIDASI
   */

  sheet
    .getRange("A:K")
    .clearDataValidations();

  /*
   * VALIDASI STATUS
   *
   * HANYA J2:J
   */

  const numberOfRows =
    Math.max(
      sheet.getMaxRows() - 1,
      1
    );

  const statusRange =
    sheet.getRange(
      2,
      10,
      numberOfRows,
      1
    );

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        STATUS_OPTIONS,
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        "Pilih: Menunggu Konfirmasi, Dikonfirmasi, Selesai, atau Dibatalkan."
      )
      .build();

  statusRange.setDataValidation(
    rule
  );

  /*
   * LEBAR KOLOM
   */

  const widths = [
    130, // A Kode
    180, // B Nama
    140, // C WhatsApp
    300, // D Layanan
    140, // E Barber
    130, // F Tanggal
    90,  // G Jam
    130, // H Total
    250, // I Catatan
    180, // J Status
    170  // K Dibuat
  ];

  widths.forEach(
    function(width, index) {

      sheet.setColumnWidth(
        index + 1,
        width
      );

    }
  );

  /*
   * HEADER BACKGROUND
   * Menggunakan format default Sheets,
   * tanpa mengubah warna secara paksa.
   */

  SpreadsheetApp.flush();
}

/* =========================
   PARSE POST DATA
========================= */

function parsePostData_(e) {

  if (!e) {

    throw new Error(
      "Data reservasi tidak ditemukan."
    );

  }

  let raw = "";

  /*
   * PRIORITAS parameter data
   */

  if (
    e.parameter &&
    e.parameter.data
  ) {

    raw =
      e.parameter.data;

  }

  /*
   * Jika tidak ada,
   * gunakan postData.contents
   */

  else if (
    e.postData &&
    e.postData.contents
  ) {

    raw =
      e.postData.contents;

  }

  if (!raw) {

    throw new Error(
      "Data reservasi tidak ditemukan."
    );

  }

  try {

    const parsed =
      JSON.parse(raw);

    if (
      !parsed ||
      typeof parsed !== "object"
    ) {

      throw new Error();

    }

    return parsed;

  } catch (_) {

    throw new Error(
      "Format data reservasi tidak valid."
    );

  }
}

/* =========================
   VALIDASI BOOKING
========================= */

function validateBooking_(data) {

  if (
    !data ||
    typeof data !== "object"
  ) {

    throw new Error(
      "Format data reservasi tidak valid."
    );

  }

  /* ---------- KODE ---------- */

  data.code =
    safeString_(
      data.code
    ).toUpperCase();

  if (
    !/^KNR-[A-Z0-9]{5,10}$/.test(
      data.code
    )
  ) {

    throw new Error(
      "Kode reservasi tidak valid."
    );

  }

  /* ---------- NAMA ---------- */

  data.name =
    safeString_(
      data.name
    );

  if (
    data.name.length < 3 ||
    data.name.length > 100
  ) {

    throw new Error(
      "Nama pelanggan harus 3–100 karakter."
    );

  }

  /* ---------- PHONE ---------- */

  data.phone =
    normalizePhone_(
      data.phone
    );

  if (
    data.phone.length < 9 ||
    data.phone.length > 15
  ) {

    throw new Error(
      "Nomor WhatsApp tidak valid."
    );

  }

  /* ---------- BARBER ---------- */

  data.barberId =
    normalizeBarberId_(
      data.barberId
    );

  /*
   * Jika ID tidak dikenali,
   * coba berdasarkan nama barber.
   */

  if (
    !BARBERS[data.barberId]
  ) {

    const barberNameInput =
      safeString_(
        data.barberName
      ).toLowerCase();

    if (
      barberNameInput ===
      "bang jaka"
    ) {

      data.barberId = "bj";

    }

    if (
      barberNameInput ===
      "kang rudi"
    ) {

      data.barberId = "kr";

    }

  }

  if (
    !BARBERS[data.barberId]
  ) {

    throw new Error(
      "Barber belum dipilih."
    );

  }

  data.barberName =
    BARBERS[data.barberId];

  /* ---------- TANGGAL ---------- */

  if (
    !isValidDate_(
      data.date
    )
  ) {

    throw new Error(
      "Tanggal reservasi tidak valid."
    );

  }

  /* ---------- JAM ---------- */

  if (
    !isValidTime_(
      data.time
    )
  ) {

    throw new Error(
      "Jam reservasi tidak valid."
    );

  }

  /* ---------- LAYANAN ---------- */

  data.services =
    Array.isArray(
      data.services
    )

      ? data.services
          .map(function(service) {
            return safeString_(
              service
            );
          })
          .filter(Boolean)
          .slice(0, 10)

      : [];

  if (
    !data.services.length
  ) {

    throw new Error(
      "Layanan belum dipilih."
    );

  }

  /* ---------- CATATAN ---------- */

  data.notes =
    safeString_(
      data.notes
    ).slice(0, 500);

  /* ---------- TOTAL ---------- */

  data.total =
    Number(data.total);

  if (
    !Number.isFinite(
      data.total
    ) ||
    data.total < 0 ||
    data.total > 100000000
  ) {

    throw new Error(
      "Total reservasi tidak valid."
    );

  }
}

/* =========================
   NORMALISASI PHONE
========================= */

function normalizePhone_(value) {

  let d =
    safeString_(value)
      .replace(/\D/g, "");

  if (!d) {
    return "";
  }

  /*
   * 08xxxxxxxxxx
   * menjadi
   * 628xxxxxxxxxx
   */

  if (
    d.indexOf("0") === 0
  ) {

    return "62" +
      d.slice(1);

  }

  /*
   * Sudah 62
   */

  if (
    d.indexOf("62") === 0
  ) {

    return d;

  }

  /*
   * Selain itu,
   * simpan angka apa adanya.
   */

  return d;
}

/* =========================
   NORMALISASI BARBER ID
========================= */

function normalizeBarberId_(id) {

  const s =
    safeString_(id)
      .toLowerCase();

  return (
    BARBER_ALIASES[s] ||
    s
  );
}

/* =========================
   SAFE STRING
========================= */

function safeString_(value) {

  return value == null
    ? ""
    : String(value).trim();

}

/* =========================
   NORMALISASI ANGKA
========================= */

function normalizeNumber_(value) {

  if (
    typeof value === "number" &&
    Number.isFinite(value)
  ) {

    return value;

  }

  const s =
    safeString_(value)
      .replace(
        /[^\d.-]/g,
        ""
      );

  const n =
    Number(s);

  return Number.isFinite(n)
    ? n
    : 0;
}

/* =========================
   NORMALISASI CREATED AT
========================= */

function normalizeCreatedAt_(value) {

  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]" &&
    !isNaN(value)
  ) {

    return value;

  }

  const s =
    safeString_(value);

  if (!s) {

    return new Date();

  }

  const d =
    new Date(s);

  return isNaN(d)
    ? new Date()
    : d;
}

/* =========================
   CEK STATUS
========================= */

function isStatus_(value) {

  const s =
    safeString_(value)
      .toLowerCase();

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

/* =========================
   NORMALISASI STATUS
========================= */

function normalizeStatus_(value) {

  const s =
    safeString_(value)
      .toLowerCase();

  if (
    s === "dikonfirmasi"
  ) {

    return "Dikonfirmasi";

  }

  if (
    s === "selesai"
  ) {

    return "Selesai";

  }

  if (
    s === "dibatalkan" ||
    s === "cancelled" ||
    s === "canceled"
  ) {

    return "Dibatalkan";

  }

  return "Menunggu Konfirmasi";
}

/* =========================
   CEK STATUS CANCEL
========================= */

function isCancelledStatus_(status) {

  return (
    normalizeStatus_(
      status
    ) === "Dibatalkan"
  );

}

/* =========================
   VALIDASI TANGGAL
========================= */

function isValidDate_(value) {

  const s =
    safeString_(value);

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(s)
  ) {

    return false;

  }

  const p =
    s.split("-")
      .map(Number);

  const dt =
    new Date(
      p[0],
      p[1] - 1,
      p[2]
    );

  return (

    dt.getFullYear() === p[0] &&

    dt.getMonth() ===
      p[1] - 1 &&

    dt.getDate() ===
      p[2]

  );
}

/* =========================
   VALIDASI JAM
========================= */

function isValidTime_(value) {

  const s =
    safeString_(value);

  if (
    !/^\d{2}:\d{2}$/.test(s)
  ) {

    return false;

  }

  const p =
    s.split(":")
      .map(Number);

  return (

    p[0] >= 0 &&
    p[0] <= 23 &&
    p[1] >= 0 &&
    p[1] <= 59

  );
}

/* =========================
   CEK DATE LIKE
========================= */

function isDateLike_(value) {

  if (

    Object.prototype.toString.call(
      value
    ) === "[object Date]" &&

    !isNaN(value)

  ) {

    return true;

  }

  const s =
    safeString_(value);

  return (

    /^\d{4}-\d{2}-\d{2}$/.test(s) ||

    /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s)

  );
}

/* =========================
   CEK TIME LIKE
========================= */

function isTimeLike_(value) {

  const s =
    safeString_(value);

  return /^\d{1,2}:\d{2}/.test(s);

}

/* =========================
   NORMALISASI TANGGAL SHEET
========================= */

function normalizeSheetDate_(value) {

  /*
   * Jika Date object
   */

  if (

    Object.prototype.toString.call(
      value
    ) === "[object Date]" &&

    !isNaN(value)

  ) {

    return Utilities.formatDate(
      value,
      TIMEZONE,
      "yyyy-MM-dd"
    );

  }

  const s =
    safeString_(value);

  /*
   * Sudah YYYY-MM-DD
   */

  if (
    /^\d{4}-\d{2}-\d{2}$/.test(s)
  ) {

    return s;

  }

  /*
   * DD/MM/YYYY atau DD-MM-YYYY
   */

  const m =
    s.match(
      /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
    );

  if (m) {

    const d =
      Number(m[1]);

    const mo =
      Number(m[2]);

    const y =
      Number(m[3]);

    const result =
      y +
      "-" +
      String(mo).padStart(2, "0") +
      "-" +
      String(d).padStart(2, "0");

    if (
      isValidDate_(result)
    ) {

      return result;

    }

  }

  /*
   * Coba Date parser
   */

  const d =
    new Date(s);

  return isNaN(d)

    ? s

    : Utilities.formatDate(
        d,
        TIMEZONE,
        "yyyy-MM-dd"
      );
}

/* =========================
   NORMALISASI JAM SHEET
========================= */

function normalizeSheetTime_(value) {

  /*
   * Jika Date object
   */

  if (

    Object.prototype.toString.call(
      value
    ) === "[object Date]" &&

    !isNaN(value)

  ) {

    return Utilities.formatDate(
      value,
      TIMEZONE,
      "HH:mm"
    );

  }

  const s =
    safeString_(value);

  const m =
    s.match(
      /^(\d{1,2}):(\d{2})/
    );

  if (!m) {

    return s;

  }

  return (

    String(
      Number(m[1])
    ).padStart(2, "0") +

    ":" +

    m[2]

  );
}

/* =========================
   GET SLOTS
========================= */

function getSlots_(
  barberId,
  date
) {

  barberId =
    normalizeBarberId_(
      barberId
    );

  date =
    safeString_(date);

  /*
   * Validasi
   */

  if (
    !BARBERS[barberId] ||
    !isValidDate_(date)
  ) {

    return {
      success: true,
      slots: []
    };

  }

  const sheet =
    getSheet_();

  /*
   * TIDAK menggunakan ensureSchema_().
   */

  return {

    success: true,

    slots:
      getBookedSlots_(
        sheet,
        barberId,
        date
      )

  };
}

/* =========================
   GET BOOKED SLOTS
========================= */

function getBookedSlots_(
  sheet,
  barberId,
  date
) {

  const last =
    sheet.getLastRow();

  if (last < 2) {

    return [];

  }

  const normalizedId =
    normalizeBarberId_(
      barberId
    );

  const barberName =
    BARBERS[
      normalizedId
    ] || "";

  /*
   * Ambil A:K
   */

  const rows =
    sheet
      .getRange(
        2,
        1,
        last - 1,
        HEADERS.length
      )
      .getValues();

  const out = [];

  rows.forEach(
    function(r) {

      const storedBarber =
        safeString_(
          r[4]
        );

      const storedDate =
        normalizeSheetDate_(
          r[5]
        );

      const storedTime =
        normalizeSheetTime_(
          r[6]
        );

      const status =
        safeString_(
          r[9]
        );

      /*
       * Hanya jam yang belum dibatalkan
       */

      if (

        storedBarber ===
          barberName &&

        storedDate ===
          date &&

        storedTime &&

        !isCancelledStatus_(
          status
        )

      ) {

        out.push(
          storedTime
        );

      }

    }
  );

  /*
   * Hapus duplikat
   */

  return Array
    .from(
      new Set(out)
    )
    .sort();

}

/* =========================
   LOOKUP RESERVASI
========================= */

function lookupBookingByCode_(
  code
) {

  const clean =
    safeString_(code)
      .toUpperCase();

  /*
   * Validasi kode
   */

  if (
    !/^KNR-[A-Z0-9]{5,10}$/.test(
      clean
    )
  ) {

    return {
      success: true,
      bookings: []
    };

  }

  const sheet =
    getSheet_();

  /*
   * TIDAK menggunakan ensureSchema_().
   */

  const last =
    sheet.getLastRow();

  if (last < 2) {

    return {
      success: true,
      bookings: []
    };

  }

  /*
   * Ambil data A:K
   */

  const rows =
    sheet
      .getRange(
        2,
        1,
        last - 1,
        HEADERS.length
      )
      .getValues();

  const bookings = [];

  rows.forEach(
    function(r) {

      if (
        safeString_(r[0])
          .toUpperCase() ===
        clean
      ) {

        bookings.push(
          makeBookingObject_(r)
        );

      }

    }
  );

  /*
   * Urutkan terbaru
   */

  bookings.sort(
    function(a, b) {

      return String(
        b.createdAt || ""
      ).localeCompare(
        String(
          a.createdAt || ""
        )
      );

    }
  );

  return {

    success: true,

    bookings:
      bookings

  };
}

/* =========================
   BUAT OBJECT BOOKING
========================= */

function makeBookingObject_(r) {

  const created =
    r[10] instanceof Date

      ? r[10].toISOString()

      : safeString_(
          r[10]
        );

  return {

    code:
      safeString_(
        r[0]
      ),

    name:
      safeString_(
        r[1]
      ),

    phone:
      normalizePhone_(
        r[2]
      ),

    services:
      safeString_(
        r[3]
      )
        .split(",")
        .map(
          function(service) {
            return service.trim();
          }
        )
        .filter(Boolean),

    barberName:
      safeString_(
        r[4]
      ),

    date:
      normalizeSheetDate_(
        r[5]
      ),

    dateLabel:
      formatDateLabel_(
        r[5]
      ),

    time:
      normalizeSheetTime_(
        r[6]
      ),

    total:
      Number(r[7]) || 0,

    notes:
      safeString_(
        r[8]
      ),

    status:
      normalizeStatus_(
        r[9]
      ),

    createdAt:
      created

  };
}

/* =========================
   FORMAT TANGGAL LABEL
========================= */

function formatDateLabel_(value) {

  const normalized =
    normalizeSheetDate_(
      value
    );

  if (
    !isValidDate_(
      normalized
    )
  ) {

    return safeString_(
      value
    );

  }

  const p =
    normalized
      .split("-")
      .map(Number);

  const names = [

    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember"

  ];

  return (

    p[2] +
    " " +
    names[p[1] - 1] +
    " " +
    p[0]

  );
}

/* =========================
   CEK KODE SUDAH ADA
========================= */

function codeExists_(
  sheet,
  code
) {

  const last =
    sheet.getLastRow();

  if (last < 2) {

    return false;

  }

  const target =
    safeString_(code)
      .toUpperCase();

  return sheet

    .getRange(
      2,
      1,
      last - 1,
      1
    )

    .getValues()

    .some(
      function(r) {

        return (
          safeString_(r[0])
            .toUpperCase() ===
          target
        );

      }
    );
}

/* =========================
   JSON RESPONSE
========================= */

function jsonOutput_(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}

/* =========================
   ERROR MESSAGE
========================= */

function getErrorMessage_(err) {

  return (

    err &&
    err.message

  )

    ? err.message

    : "Terjadi kesalahan server.";

}

/* =========================
   TEST CONNECTION
========================= */

/*
 * Bisa dijalankan manual
 * dari Google Apps Script.
 */

function testConnection() {

  const sheet =
    getSheet_();

  normalizeSheetStructure_(
    sheet
  );

  configureSheet_(
    sheet
  );

  return {

    success: true,

    message:
      "Koneksi Google Sheets berhasil.",

    sheet:
      SHEET_NAME,

    columns:
      HEADERS.length

  };

}
