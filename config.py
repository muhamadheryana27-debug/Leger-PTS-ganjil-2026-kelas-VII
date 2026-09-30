"""Konfigurasi tunggal untuk seluruh projek raport PTS.

Satu-satunya tempat untuk mengubah:
- path template / database / output (dua folder: individu & gabungan)
- TITI MANGSA (TEMPAT_RAPOT + TANGGAL_RAPOT) — cukup tulis di sini
- KKM
- daftar mapel
- slot paragraf nilai (MAPEL_SLOT) — nomor paragraf tetap per template

Template Print_Rapot.docx (193 paragraf) memakai dua sistem:
1. Placeholder UNIK -> diisi by pencarian teks (tahan geser paragraf):
   [Kelas], [Nama], [NIS/NISN], [jumlah], [NAMA_WALI], [NIP_WALI]
2. Placeholder GENERIK berulang ([angka]/[huruf]/[deskripsi]) -> HARUS
   diisi by nomor paragraf via MAPEL_SLOT di bawah.
   Jika template diedit di Word (tambah/geser paragraf), cek ulang dengan:
   unzip -p Print_Rapot.docx word/document.xml | grep -o ...
   atau jalankan validasi otomatis DocxTemplate.validate().
"""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_DOCX = str(BASE_DIR / "Print_Rapot.docx")
DATABASE_XLSX = str(BASE_DIR / "database.xlsx")
DATABASE_ODS = str(BASE_DIR / "database.ods")
# Database utama (isi xlsx & ods identik; --database bisa menimpa)
DATABASE_DEFAULT = DATABASE_XLSX

# Folder 1: file individu per siswa (hasil cetak_rapot.py)
OUTPUT_DIR = str(BASE_DIR / "hasil_rapot_docx")
OUTPUT_INDIVIDU_DIR = OUTPUT_DIR  # alias, nama lebih jelas
# Folder 2: file gabungan per kelas (hasil gabung_rapot.py, dijalankan setelah cetak)
OUTPUT_GABUNGAN_DIR = str(BASE_DIR / "hasil_rapot_gabungan")

# ── TITI MANGSA ─────────────────────────────────────────────────────────────
# Cukup ubah di sini. Format YYYY-MM-DD, contoh: "2026-10-20".
# Jika None, dipakai tanggal hari ini saat program dijalankan.
# Opsi CLI --tanggal selalu menang atas nilai ini.
TEMPAT_RAPOT = "Wanayasa"
TANGGAL_RAPOT = None  # contoh: "2026-10-20"

# Kriteria Ketuntasan Minimal — dipakai cetak_rapot.py & leger_interaktif.py
KKM = 71

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

# Urutan kanonis mapel bernilai (ada di database). Lookup memakai NAMA kolom,
# jadi urutan kolom di database boleh beda tanpa merusak hasil cetak.
# Keterampilan TIDAK masuk sini (tidak ada kolom nilainya -> diisi '-').
DAFTAR_MAPEL = [
    'PAI', 'PPKn', 'B.IND', 'B.ING', 'MTK', 'IPA',
    'IPS', 'PKY', 'PJOK', 'INF', 'B.SUN',
]

# Slot nilai: mapel -> nomor paragraf untuk tiap kolom.
# Struktur rapor: 1-9 triple (angka,huruf,deskripsi), 10a Keterampilan
# (angka saja -> KET_ANGKA), 10b Informatika triple, 11a B.Sunda
# (angka di baris header Mulok p117, huruf+deskripsi di barisnya).
# Cek ulang jika template berubah!
MAPEL_SLOT = {
    'PAI':   {'angka': 42, 'huruf': 43, 'deskripsi': 44},
    'PPKn':  {'angka': 48, 'huruf': 49, 'deskripsi': 50},
    'B.IND': {'angka': 54, 'huruf': 55, 'deskripsi': 56},
    'B.ING': {'angka': 60, 'huruf': 61, 'deskripsi': 62},
    'MTK':   {'angka': 66, 'huruf': 67, 'deskripsi': 68},
    'IPA':   {'angka': 72, 'huruf': 73, 'deskripsi': 74},
    'IPS':   {'angka': 78, 'huruf': 79, 'deskripsi': 80},
    'PKY':   {'angka': 84, 'huruf': 85, 'deskripsi': 86},
    'PJOK':  {'angka': 90, 'huruf': 91, 'deskripsi': 92},
    'INF':   {'angka': 111, 'huruf': 112, 'deskripsi': 113},
    'B.SUN': {'angka': 117, 'huruf': 127, 'deskripsi': 128},
}

# Baris 10a Keterampilan & 11b BTQ di template semuanya '-' (literal,
# tanpa placeholder) -> tidak ada yang perlu diisi program.

# Baris Jumlah: [jumlah] unik (by pencarian), [huruf] generik (by indeks)
TOTAL_HURUF = 137

# Placeholder UNIK di template (diisi by pencarian teks)
PH_KELAS = "[Kelas]"
PH_NAMA = "[Nama]"
PH_NIS = "[NIS/NISN]"
PH_JUMLAH = "[jumlah]"
PH_TANGGAL = "[TITI_MANGSA]"
PH_NAMA_WALI = "[NAMA_WALI]"
PH_NIP_WALI = "[NIP_WALI]"

# Marker generik yang HARUS ada di tiap slot (untuk validasi template)
MARKER_ANGKA = "[angka]"
MARKER_HURUF = "[huruf]"
MARKER_DESKRIPSI = "[deskripsi]"

# Namespace WordprocessingML (+ xml:space untuk preservasi spasi)
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = 'http://www.w3.org/XML/1998/namespace'

# Kolom wajib di database
KOLOM_WAJIB = ['Nama', 'Kelas', 'NIS/NISN', 'Wali Kelas', 'NIP'] + DAFTAR_MAPEL
