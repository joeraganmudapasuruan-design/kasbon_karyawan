# Kasbon Karyawan — Skrip Google Sheets

Otomatisasi untuk file **KASBON KARYAWAN** (Auto Plus Carwash & Detailing).

File: **`KasbonKaryawan.gs`** — tempel ke Apps Script (Ekstensi > Apps Script).

## Fitur
1. **Tanggal otomatis** — isi Nama di `MASTER INPUT` → kolom *Tanggal* terisi tanggal hari ini (nilai tetap). Sama untuk *Tgl Terima* di `JAMINAN`.
2. **Kalender (date picker)** — kolom *Tanggal* & *Tgl Terima* bisa diklik untuk pilih tanggal.
3. **Jaminan merah** — kolom *Jaminan* di `MASTER INPUT` jadi merah bila ada **Kasbon ≥ Rp 1.000.000** yang jaminannya kosong / "-" / "Tidak Ada".
4. **Kode karyawan otomatis** — isi nama baru di `MASTER KARYAWAN` → kolom *Kode* terisi `K007`, `K008`, … (nilai tetap). Kode lama `K001`–`K006` tidak diubah.

## Pasang (sekali saja)
1. Buka Sheet → **Ekstensi > Apps Script**.
2. Hapus isi editor, tempel seluruh `KasbonKaryawan.gs`, **Save**.
3. **Refresh** Sheet → muncul menu **Kasbon** di atas.
4. Klik **Kasbon > Pasang semua (sekali jalan)** → beri izin (Review permissions > Allow).

> Catatan: pastikan zona waktu project Apps Script = **Asia/Jakarta (WIB)** agar "hari ini" tepat (Project Settings > Time zone).
