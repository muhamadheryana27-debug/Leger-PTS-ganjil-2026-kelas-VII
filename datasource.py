"""Akses database (xlsx/ods) — baca + validasi + normalisasi."""

import pandas as pd

from config import DAFTAR_MAPEL


def _engine_untuk(path):
    """Pilih engine pandas berdasar ekstensi: xlsx->openpyxl, ods->odf."""
    return 'openpyxl' if str(path).lower().endswith('.xlsx') else 'odf'


def load_database(path, kelas_filter=None):
    """Baca database -> DataFrame bersih.

    - dtype NIP/NIS/NISN/Kelas dipaksa str agar '19720303...' & '25267001/...'
      tidak berubah jadi float/pecah.
    - Validasi kolom wajib, raise ValueError yang jelas jika kurang.
    - Normalisasi: strip Kelas, drop baris tanpa Nama (dicatat jumlahnya).
    - Opsional filter satu kelas (untuk --kelas).

    Returns:
        (df, info: dict{jml_awal, dibuang_tanpa_nama, kolom_hilang})
    """
    df = pd.read_excel(
        path, engine=_engine_untuk(path),
        dtype={'NIP': str, 'NIS/NISN': str, 'Kelas': str},
    )
    jml_awal = len(df)

    kolom_hilang = [c for c in ['Nama', 'Kelas'] if c not in df.columns]
    if kolom_hilang:
        raise ValueError(f"Kolom wajib hilang di {path}: {kolom_hilang}")

    mapel_hilang = [m for m in DAFTAR_MAPEL if m not in df.columns]
    if mapel_hilang:
        # Tidak fatal (siswa tetap dicetak, mapel tsb jadi '-'), tapi beri tahu.
        print(f"Peringatan: kolom mapel tidak ada di database: {mapel_hilang}")

    if 'Kelas' in df.columns:
        df['Kelas'] = df['Kelas'].astype(str).str.strip()

    sebelum = len(df)
    df = df.dropna(subset=['Nama'])
    # dropna tidak menangkap string kosong
    df = df[df['Nama'].astype(str).str.strip() != '']
    dibuang = sebelum - len(df)

    if kelas_filter:
        df = df[df['Kelas'] == kelas_filter].copy()

    df = df.reset_index(drop=True)
    return df, {
        "jml_awal": jml_awal,
        "dibuang_tanpa_nama": int(dibuang),
        "mapel_hilang": mapel_hilang,
    }
