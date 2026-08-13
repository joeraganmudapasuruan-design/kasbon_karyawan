/**
 * OTOMATISASI KASBON KARYAWAN (Google Sheets)
 * Auto Plus Carwash & Detailing
 *
 * Fitur:
 *   1. TANGGAL OTOMATIS  -> isi Nama di MASTER INPUT, kolom "Tanggal" langsung
 *      terisi tanggal hari ini (nilai asli, bukan rumus). Juga di JAMINAN
 *      untuk kolom "Tgl Terima".
 *   2. KALENDER (DATE PICKER) -> kolom Tanggal & Tgl Terima bisa diklik untuk
 *      memilih tanggal dari kalender.
 *   3. JAMINAN MERAH -> kolom "Jaminan" di MASTER INPUT jadi merah kalau ada
 *      Kasbon >= Rp 1.000.000 yang jaminannya kosong / "-" / "Tidak Ada".
 *   4. KODE KARYAWAN OTOMATIS -> isi Nama karyawan baru di MASTER KARYAWAN,
 *      kolom "Kode" otomatis terisi K007, K008, ... (nilai tetap). Kode lama
 *      (K001-K006) TIDAK disentuh.
 *
 * ============================================================
 * CARA PASANG (sekali saja):
 *   1. Buka Google Sheet-nya
 *   2. Menu  Ekstensi (Extensions)  >  Apps Script
 *   3. Hapus semua isi editor, tempel SELURUH kode ini, klik Save (disket)
 *   4. Muat ulang (refresh) Sheet-nya. Muncul menu baru "Kasbon" di atas.
 *   5. Klik menu  Kasbon > Pasang semua (sekali jalan).
 *   6. Muncul permintaan izin sekali: Review permissions > pilih akun >
 *      Advanced > Allow.  Selesai.
 * ============================================================
 */

var FORMAT_TANGGAL  = 'dd mmmm yyyy';
var BATAS_JAMINAN   = 1000000;      // kasbon mulai nominal ini wajib ada jaminan
var WARNA_MERAH_BG  = '#FFC7CE';
var WARNA_MERAH_TXT = '#9C0006';

/* Konfigurasi sheet yang punya kolom tanggal-otomatis + kalender. */
var KONFIG = [
  {
    sheet: 'MASTER INPUT',
    barisJudul: 5,
    judulPemicu:  'Nama Karyawan',    // dicari di baris judul, bukan huruf kolom
    judulTanggal: 'Tanggal',
    barisMulai: 6,
    barisAkhir: 500,
    // untuk aturan "Jaminan merah":
    pasangMerah:   true,
    judulJenis:    'Jenis Transaksi',
    judulNominal:  'Nominal (Rp)',
    judulJaminan:  'Jaminan'
  },
  {
    sheet: 'JAMINAN',
    barisJudul: 5,
    judulPemicu:  'Nama Karyawan',
    judulTanggal: 'Tgl Terima',
    barisMulai: 6,
    barisAkhir: 55,
    pasangMerah: false
  }
];

/* Konfigurasi kode-karyawan-otomatis. */
var KONFIG_KODE = {
  sheet: 'MASTER KARYAWAN',
  barisJudul: 4,
  judulPemicu: 'Nama Karyawan',
  judulKode:   'Kode',
  prefiks: 'K',
  digit: 3,          // K001, K002, ...
  barisMulai: 5,
  barisAkhir: 54
};


/* ============================================================
 * MENU
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kasbon')
    .addItem('Pasang semua (sekali jalan)', 'pasangSemua')
    .addToUi();
}


/* ============================================================
 * UTILITAS
 * ============================================================ */

/** Cari nomor kolom berdasarkan judulnya. 0 kalau tidak ketemu. */
function cariKolom(sh, barisJudul, judul) {
  var judulSemua = sh.getRange(barisJudul, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < judulSemua.length; i++) {
    if (String(judulSemua[i]).trim() === judul) return i + 1;
  }
  return 0;
}

/** Nomor kolom -> huruf kolom (1 -> A, 27 -> AA). */
function hurufKolom(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "K" + angka dipadatkan jadi K007 dst. */
function buatKode(prefiks, num, digit) {
  var s = String(num);
  while (s.length < digit) s = '0' + s;
  return prefiks + s;
}

/** Nomor kode tertinggi yang sudah ada di kolom Kode (0 kalau kosong). */
function kodeTertinggi(sh, cfg, kolKode) {
  var n = cfg.barisAkhir - cfg.barisMulai + 1;
  var vals = sh.getRange(cfg.barisMulai, kolKode, n, 1).getValues();
  var max = 0;
  for (var i = 0; i < vals.length; i++) {
    var m = String(vals[i][0]).match(/(\d+)/);
    if (m) { var num = parseInt(m[1], 10); if (num > max) max = num; }
  }
  return max;
}


/* ============================================================
 * SETUP SEKALI JALAN
 * ============================================================ */
function pasangSemua() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var laporan = [];

  // ---- 1 & 2: bersihkan rumus tanggal + pasang kalender + format ----
  for (var i = 0; i < KONFIG.length; i++) {
    var cfg = KONFIG[i];
    var sh = ss.getSheetByName(cfg.sheet);
    if (!sh) { laporan.push('- Sheet "' + cfg.sheet + '" tidak ada, dilewati.'); continue; }

    var kolNama = cariKolom(sh, cfg.barisJudul, cfg.judulPemicu);
    var kolTgl  = cariKolom(sh, cfg.barisJudul, cfg.judulTanggal);
    if (!kolNama || !kolTgl) {
      laporan.push('- ' + cfg.sheet + ': kolom "' + cfg.judulPemicu + '" / "' +
                   cfg.judulTanggal + '" tidak ketemu. Dilewati.');
      continue;
    }

    var n = cfg.barisAkhir - cfg.barisMulai + 1;
    var sel   = sh.getRange(cfg.barisMulai, kolTgl, n, 1);
    var nama  = sh.getRange(cfg.barisMulai, kolNama, n, 1).getValues();
    var isi   = sel.getValues();
    var rumus = sel.getFormulas();

    var dibersihkan = 0;
    for (var j = 0; j < n; j++) {
      var kosongNama = String(nama[j][0]).trim() === '';
      var adaRumus   = String(rumus[j][0]).charAt(0) === '=';
      if (kosongNama || adaRumus) {
        if (isi[j][0] !== '' || adaRumus) { isi[j][0] = ''; dibersihkan++; }
      }
    }
    sel.setValues(isi);
    sel.setNumberFormat(FORMAT_TANGGAL);

    // kalender / date picker
    var aturanTgl = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .setHelpText('Klik untuk memilih tanggal dari kalender.')
      .build();
    sel.setDataValidation(aturanTgl);

    laporan.push('- ' + cfg.sheet + ': kalender terpasang di "' + cfg.judulTanggal +
                 '", ' + dibersihkan + ' sel dirapikan.');

    // ---- 3: warnai Jaminan merah langsung (khusus yang pasangMerah) ----
    if (cfg.pasangMerah) {
      var pesan = warnaiSemuaJaminan(sh, cfg);
      laporan.push('- ' + cfg.sheet + ': ' + pesan);
    }
  }

  // ---- 4: cek kolom Kode di MASTER KARYAWAN ----
  var shk = ss.getSheetByName(KONFIG_KODE.sheet);
  if (shk) {
    var kolKode = cariKolom(shk, KONFIG_KODE.barisJudul, KONFIG_KODE.judulKode);
    var kolN    = cariKolom(shk, KONFIG_KODE.barisJudul, KONFIG_KODE.judulPemicu);
    if (kolKode && kolN) {
      laporan.push('- MASTER KARYAWAN: kode otomatis aktif (lanjut dari K' +
                   buatKode('', kodeTertinggi(shk, KONFIG_KODE, kolKode) + 1, KONFIG_KODE.digit) +
                   '). Kode lama tidak diubah.');
    } else {
      laporan.push('- MASTER KARYAWAN: kolom "Kode"/"Nama Karyawan" tidak ketemu.');
    }
  }

  SpreadsheetApp.getUi().alert('Pemasangan selesai\n\n' + laporan.join('\n') +
                               '\n\nSemua otomatis sekarang aktif.');
}

/** TRUE bila baris ini melanggar: Kasbon >= batas tapi jaminan kosong/-/Tidak Ada. */
function langgarJaminan(jenis, nominal, jaminan) {
  var j = String(jaminan).trim();
  var kosong = (j === '' || j === '-' || j === 'Tidak Ada');
  var nom = Number(nominal);
  return String(jenis).trim() === 'Kasbon' && !isNaN(nom) && nom >= BATAS_JAMINAN && kosong;
}

/** Buang aturan pemformatan bersyarat lama yang menyentuh kolom Jaminan. */
function buangCFJaminan(sh, kolJam) {
  var lama = sh.getConditionalFormatRules();
  var simpan = [];
  for (var i = 0; i < lama.length; i++) {
    var rgs = lama[i].getRanges();
    var sentuh = false;
    for (var k = 0; k < rgs.length; k++) {
      if (rgs[k].getColumn() <= kolJam && rgs[k].getLastColumn() >= kolJam) { sentuh = true; break; }
    }
    if (!sentuh) simpan.push(lama[i]);
  }
  sh.setConditionalFormatRules(simpan);
}

/**
 * Warnai LANGSUNG kolom Jaminan (bukan lewat pemformatan bersyarat, biar pasti
 * jalan). Merah = Kasbon >= batas tanpa jaminan. Dipakai saat setup (semua baris).
 */
function warnaiSemuaJaminan(sh, cfg) {
  var kolJns = cariKolom(sh, cfg.barisJudul, cfg.judulJenis);
  var kolNom = cariKolom(sh, cfg.barisJudul, cfg.judulNominal);
  var kolJam = cariKolom(sh, cfg.barisJudul, cfg.judulJaminan);
  if (!kolJns || !kolNom || !kolJam) {
    return 'pewarnaan DILEWATI (ada kolom yang tidak ketemu).';
  }

  buangCFJaminan(sh, kolJam);   // hapus aturan lama yang bermasalah

  var n   = cfg.barisAkhir - cfg.barisMulai + 1;
  var jns = sh.getRange(cfg.barisMulai, kolJns, n, 1).getValues();
  var nom = sh.getRange(cfg.barisMulai, kolNom, n, 1).getValues();
  var sel = sh.getRange(cfg.barisMulai, kolJam, n, 1);
  var jam = sel.getValues();
  var bg  = sel.getBackgrounds();
  var fc  = sel.getFontColors();

  var jml = 0;
  for (var i = 0; i < n; i++) {
    if (langgarJaminan(jns[i][0], nom[i][0], jam[i][0])) {
      bg[i][0] = WARNA_MERAH_BG; fc[i][0] = WARNA_MERAH_TXT; jml++;
    } else {
      bg[i][0] = null; fc[i][0] = null;   // kembalikan ke normal
    }
  }
  sel.setBackgrounds(bg);
  sel.setFontColors(fc);

  return jml + ' baris Kasbon >= Rp ' + BATAS_JAMINAN.toLocaleString('id-ID') +
         ' tanpa jaminan diberi warna merah.';
}

/** Warnai ulang kolom Jaminan hanya untuk baris r1..r2 (dipakai saat onEdit). */
function warnaiJaminanBaris(sh, cfg, r1, r2) {
  var kolJns = cariKolom(sh, cfg.barisJudul, cfg.judulJenis);
  var kolNom = cariKolom(sh, cfg.barisJudul, cfg.judulNominal);
  var kolJam = cariKolom(sh, cfg.barisJudul, cfg.judulJaminan);
  if (!kolJns || !kolNom || !kolJam) return;

  var n   = r2 - r1 + 1;
  var jns = sh.getRange(r1, kolJns, n, 1).getValues();
  var nom = sh.getRange(r1, kolNom, n, 1).getValues();
  var sel = sh.getRange(r1, kolJam, n, 1);
  var jam = sel.getValues();
  var bg  = sel.getBackgrounds();
  var fc  = sel.getFontColors();

  for (var i = 0; i < n; i++) {
    if (langgarJaminan(jns[i][0], nom[i][0], jam[i][0])) {
      bg[i][0] = WARNA_MERAH_BG; fc[i][0] = WARNA_MERAH_TXT;
    } else {
      bg[i][0] = null; fc[i][0] = null;
    }
  }
  sel.setBackgrounds(bg);
  sel.setFontColors(fc);
}


/* ============================================================
 * PEMICU OTOMATIS SETIAP EDIT
 * ============================================================ */
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var nama = sh.getName();

  for (var i = 0; i < KONFIG.length; i++) {
    if (KONFIG[i].sheet === nama) {
      var cfg = KONFIG[i];
      stempelTanggal(e, sh, cfg);
      // warnai ulang baris yang diedit (kalau sheet ini pakai aturan merah)
      if (cfg.pasangMerah) {
        var r1 = Math.max(e.range.getRow(), cfg.barisMulai);
        var r2 = Math.min(e.range.getLastRow(), cfg.barisAkhir);
        if (r2 >= r1) warnaiJaminanBaris(sh, cfg, r1, r2);
      }
      return;
    }
  }
  if (KONFIG_KODE.sheet === nama) { stempelKode(e, sh, KONFIG_KODE); return; }
}

/** Isi/kosongkan tanggal mengikuti kolom nama. */
function stempelTanggal(e, sh, cfg) {
  var kolNama = cariKolom(sh, cfg.barisJudul, cfg.judulPemicu);
  var kolTgl  = cariKolom(sh, cfg.barisJudul, cfg.judulTanggal);
  if (!kolNama || !kolTgl) return;
  if (kolNama < e.range.getColumn() || kolNama > e.range.getLastColumn()) return;

  var r1 = Math.max(e.range.getRow(), cfg.barisMulai);
  var r2 = Math.min(e.range.getLastRow(), cfg.barisAkhir);
  if (r2 < r1) return;

  var n = r2 - r1 + 1;
  var nama = sh.getRange(r1, kolNama, n, 1).getValues();
  var sel  = sh.getRange(r1, kolTgl, n, 1);
  var tgl  = sel.getValues();

  var hariIni = new Date();
  hariIni.setHours(0, 0, 0, 0);

  var berubah = false;
  for (var j = 0; j < n; j++) {
    var adaNama = String(nama[j][0]).trim() !== '';
    var adaTgl  = tgl[j][0] !== '' && tgl[j][0] !== null;
    if (adaNama && !adaTgl)      { tgl[j][0] = hariIni; berubah = true; }
    else if (!adaNama && adaTgl) { tgl[j][0] = '';      berubah = true; }
  }
  if (!berubah) return;
  sel.setValues(tgl);
  sel.setNumberFormat(FORMAT_TANGGAL);
}

/** Isi/kosongkan Kode karyawan mengikuti kolom nama. Kode lama tidak disentuh. */
function stempelKode(e, sh, cfg) {
  var kolNama = cariKolom(sh, cfg.barisJudul, cfg.judulPemicu);
  var kolKode = cariKolom(sh, cfg.barisJudul, cfg.judulKode);
  if (!kolNama || !kolKode) return;
  if (kolNama < e.range.getColumn() || kolNama > e.range.getLastColumn()) return;

  var r1 = Math.max(e.range.getRow(), cfg.barisMulai);
  var r2 = Math.min(e.range.getLastRow(), cfg.barisAkhir);
  if (r2 < r1) return;

  var n = r2 - r1 + 1;
  var nama = sh.getRange(r1, kolNama, n, 1).getValues();
  var sel  = sh.getRange(r1, kolKode, n, 1);
  var kode = sel.getValues();

  var maxNum = kodeTertinggi(sh, cfg, kolKode);   // mis. 6 dari K006

  var berubah = false;
  for (var j = 0; j < n; j++) {
    var adaNama = String(nama[j][0]).trim() !== '';
    var adaKode = String(kode[j][0]).trim() !== '';
    if (adaNama && !adaKode) {                     // nama baru -> kode berikutnya
      maxNum++;
      kode[j][0] = buatKode(cfg.prefiks, maxNum, cfg.digit);
      berubah = true;
    } else if (!adaNama && adaKode) {              // nama dihapus -> kode dikosongkan
      kode[j][0] = '';
      berubah = true;
    }
    // nama & kode dua-duanya ada -> JANGAN disentuh (kode lama aman)
  }
  if (berubah) sel.setValues(kode);
}
