"""Fungsi utilitas murni (tanpa pandas/lxml) — mudah dites unit."""

import datetime
import math
import re

from config import BULAN_ID


def get_tempat_default():
    """Ambil TEMPAT_RAPOT dari config tanpa circular import bermasalah."""
    try:
        from config import TEMPAT_RAPOT
        return TEMPAT_RAPOT or "Wanayasa"
    except ImportError:
        return "Wanayasa"

_SATUAN = [
    '', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam',
    'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas',
]

# Karakter ilegal filename Windows + strip spasi/titik di ujung
_ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|]')


def nilai_ke_int(val):
    """Ubah nilai sel menjadi int (pembulatan .5 ke atas).

    Returns:
        (ok: bool, v: int|None)
        ok=False untuk None/NaN/''/teks non-angka.
    """
    if val is None:
        return False, None
    if isinstance(val, float) and math.isnan(val):
        return False, None
    if isinstance(val, str):
        if val.strip() == '' or val.strip().lower() == 'nan':
            return False, None
    try:
        v = int(math.floor(float(val) + 0.5))
        return True, v
    except (ValueError, TypeError):
        return False, None


def angka_ke_huruf(n):
    """Ubah angka menjadi kata Indonesia. Gagal -> '-'.

    Mendukung 0 s.d. ratusan ribu (untuk [huruf] JUMLAH total, mis. 935
    -> 'Sembilan Ratus Tiga Puluh Lima'). Di atas itu fallback str(n).

    Beda dengan versi lama: pakai floor(x+0.5) agar 70.5 -> 71
    (round() Python membulatkan 70.5 -> 70 / bankers rounding),
    dan tidak menelan exception diam-diam selain yang diharapkan.
    """
    ok, v = nilai_ke_int(n)
    if not ok:
        # Kompatibilitas: teks non-angka (mis. 'VII A' nyasar) -> '-'
        return '-'
    if v < 0:
        return str(v)
    if v >= 1000000:
        return str(v)
    return _terbilang(v)


def _terbilang(v):
    """Rekursi ejaan Indonesia untuk 0 <= v < 1.000.000."""
    if v < 12:
        return _SATUAN[v]
    if v < 20:
        return _terbilang(v - 10) + ' Belas'
    if v < 100:
        puluh = _SATUAN[v // 10] + ' Puluh'
        sisa = '' if v % 10 == 0 else ' ' + _SATUAN[v % 10]
        return (puluh + sisa).strip()
    if v < 200:
        return 'Seratus' + ('' if v == 100 else ' ' + _terbilang(v - 100))
    if v < 1000:
        return (_terbilang(v // 100) + ' Ratus'
                + ('' if v % 100 == 0 else ' ' + _terbilang(v % 100)))
    if v < 2000:
        return 'Seribu' + ('' if v == 1000 else ' ' + _terbilang(v - 1000))
    return (_terbilang(v // 1000) + ' Ribu'
            + ('' if v % 1000 == 0 else ' ' + _terbilang(v % 1000)))


def status_tuntas(nilai_int, kkm):
    """'Tuntas' jika nilai >= kkm, else 'Belum Tuntas'."""
    return "Tuntas" if nilai_int >= kkm else "Belum Tuntas"


def format_tanggal_id(dt=None, tempat=None):
    """'Wanayasa, 22 September 2026'. Default: hari ini + TEMPAT_RAPOT config."""
    dt = dt or datetime.datetime.now()
    tempat = tempat or get_tempat_default()
    return f"{tempat}, {dt.day} {BULAN_ID[dt.month]} {dt.year}"


def parse_tanggal_arg(s):
    """Parse argumen --tanggal 'YYYY-MM-DD' -> datetime. None jika kosong."""
    if not s:
        return None
    return datetime.datetime.strptime(s, "%Y-%m-%d")


def resolve_tanggal(tanggal_arg=None):
    """Prioritas: --tanggal CLI > TANGGAL_RAPOT di config.py > hari ini.

    Returns:
        (datetime, sumber: str) sumber salah satu dari
        'cli' / 'config' / 'hari-ini'.
    """
    if tanggal_arg:
        return parse_tanggal_arg(tanggal_arg), "cli"
    try:
        from config import TANGGAL_RAPOT
    except ImportError:
        TANGGAL_RAPOT = None
    if TANGGAL_RAPOT:
        return parse_tanggal_arg(TANGGAL_RAPOT), "config"
    return datetime.datetime.now(), "hari-ini"


def bersihkan_nip(nip):
    """Normalisasi NIP: hilangkan '.0', tangani nan/None/'' -> ''."""
    s = '' if nip is None else str(nip).strip()
    if s.lower() in ('', 'nan', 'none', 'nat'):
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    return s


def bersihkan_nis(nis):
    """Normalisasi NIS/NISN (perlakuan sama dengan NIP)."""
    return bersihkan_nip(nis)


def format_nip_display(nip):
    """'NIP. 123...' atau 'NIP. -' jika kosong."""
    bersih = bersihkan_nip(nip)
    return f"NIP. {bersih}" if bersih else "NIP. -"


def sanitize_nama_file(teks):
    """Buat komponen filename aman: ganti ilegal -> '-', rapikan spasi."""
    s = str(teks).strip()
    s = _ILLEGAL_CHARS.sub('-', s)
    s = re.sub(r'\s+', ' ', s).strip(' .')
    return s


def buat_nama_file(kelas, nama, terpakai=None):
    """'VIIA_ADI HERMAWAN.docx' — aman + deduplikasi (_2, _3...).

    Args:
        terpakai: set() nama file yang sudah dipakai (untuk hindari overwrite
            saat dua siswa bernama sama di kelas yang sama).
    """
    kelas_bersih = sanitize_nama_file(kelas).replace(' ', '').upper() or 'TANPAKELAS'
    nama_bersih = sanitize_nama_file(nama).upper() or 'TANPANAMA'
    dasar = f"{kelas_bersih}_{nama_bersih}"
    kandidat = dasar + ".docx"
    if terpakai is None:
        return kandidat
    i = 2
    while kandidat in terpakai:
        kandidat = f"{dasar}_{i}.docx"
        i += 1
    return kandidat
