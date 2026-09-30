"""Cetak raport PTS massal -> DOCX per siswa.

Alur resmi (dua folder terpisah):
    1. python cetak_rapot.py                  # -> hasil_rapot_docx/ (individu)
    2. python gabung_rapot.py                 # -> hasil_rapot_gabungan/ (per kelas)
    atau sekaligus:
       python cetak_rapot.py --gabung

Titimangsa cukup ditulis di config.py (TANGGAL_RAPOT, format YYYY-MM-DD).
Opsi --tanggal selalu menang atas config; jika keduanya kosong dipakai hari ini.

Contoh:
    python cetak_rapot.py --dry-run --limit 5
    python cetak_rapot.py --kelas "VII A" --dry-run
    python cetak_rapot.py --gabung --overwrite
    python cetak_rapot.py --tanggal 2026-10-20 --output hasil_rapot_docx

Refactor modular:
    config.py           konstanta (KKM, MAPEL_SLOT, indeks paragraf, TITI MANGSA)
    utils.py            angka_ke_huruf, tanggal, NIP/NIS, nama file
    template_engine.py  DocxTemplate (lxml+zip, dibuka 1x)
    datasource.py       load_database + validasi kolom
    gabung_rapot.py     gabung individu -> per kelas
"""

import argparse
import logging
import os
import sys

try:
    import pandas as pd
except ImportError:  # pesan ramah jika venv belum di-install
    pd = None

from config import DATABASE_DEFAULT, KKM, OUTPUT_GABUNGAN_DIR, OUTPUT_INDIVIDU_DIR, TEMPLATE_DOCX
from datasource import load_database
from template_engine import DocxTemplate
from utils import (
    angka_ke_huruf,  # noqa: F401  (re-ekspor agar API lama tetap ada)
    buat_nama_file,
    format_tanggal_id,
    resolve_tanggal,
)

log = logging.getLogger("cetak_rapot")

# ---------------------------------------------------------------------------
# Kompatibilitas API lama — wrapper tipis ke modul baru.
# Kode lama yang memanggil fungsi ini tetap jalan.
# ---------------------------------------------------------------------------

def set_para_text(paras, idx, text, bold=None, underline=None):
    """API lama. Gaya mengikuti template bila bold/underline=None."""
    from template_engine import _gaya_paragraf, _tulis_paragraf
    if idx >= len(paras):
        return
    p = paras[idx]
    if bold is None or underline is None:
        b_auto, u_auto = _gaya_paragraf(p)
        bold = b_auto if bold is None else bold
        underline = u_auto if underline is None else underline
    _tulis_paragraf(p, text, bold=bold, underline=underline)


def ganti_teks_placeholder(paras, placeholder, pengganti, bold=None, underline=None):
    """API lama. Gaya mengikuti template bila bold/underline=None."""
    from template_engine import _gaya_paragraf, _teks_paragraf, _tulis_paragraf
    for p in paras:
        full_text = _teks_paragraf(p)
        if placeholder in full_text:
            if bold is None or underline is None:
                b_auto, u_auto = _gaya_paragraf(p)
                b = b_auto if bold is None else bold
                u = u_auto if underline is None else underline
            else:
                b, u = bold, underline
            _tulis_paragraf(p, full_text.replace(placeholder, str(pengganti)),
                            bold=b, underline=u)


def buat_file_rapot(template_bytes, row, tgl_cetak=None, kkm=KKM):
    """API lama -> bytes DOCX. tgl_cetak opsional (dulu selalu hari ini)."""
    tmpl = DocxTemplate(template_bytes)
    data, _info = tmpl.render(row, tgl_cetak or format_tanggal_id(), kkm=kkm)
    return data


# ---------------------------------------------------------------------------
# CLI baru
# ---------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(description="Cetak raport PTS massal ke DOCX per siswa.")
    p.add_argument("--template", default=TEMPLATE_DOCX)
    p.add_argument("--database", default=DATABASE_DEFAULT)
    p.add_argument("--output", default=OUTPUT_INDIVIDU_DIR,
                   help="Folder file individu (default dari config.py)")
    p.add_argument("--kelas", default=None, help='Filter satu kelas, mis. "VII A"')
    p.add_argument("--tanggal", default=None,
                   help="Tanggal titimangsa YYYY-MM-DD (menang atas TANGGAL_RAPOT di config.py)")
    p.add_argument("--kkm", type=int, default=KKM)
    p.add_argument("--limit", type=int, default=None, help="Batasi N siswa pertama (uji coba)")
    p.add_argument("--dry-run", action="store_true", help="Proses tanpa menulis file")
    p.add_argument("--overwrite", action="store_true",
                   help="Tulis ulang file yang sudah ada (default: lewati)")
    p.add_argument("--gabung", action="store_true",
                   help="Setelah cetak, langsung gabungkan per kelas ke folder gabungan")
    p.add_argument("--folder-gabungan", default=OUTPUT_GABUNGAN_DIR,
                   help="Folder file gabungan (default dari config.py)")
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv=None):
    logging.basicConfig(
        level=logging.DEBUG if (argv and ("-v" in argv or "--verbose" in argv)) else logging.INFO,
        format="%(levelname)s: %(message)s",
    )
    args = build_parser().parse_args(argv)

    if pd is None:
        log.error("pandas belum terinstall. Jalankan: pip install -r requirements.txt")
        return 2
    if not os.path.exists(args.template):
        log.error("Template tidak ditemukan: %s", args.template)
        return 2
    if not os.path.exists(args.database):
        log.error("Database tidak ditemukan: %s", args.database)
        return 2

    try:
        dt, sumber_tgl = resolve_tanggal(args.tanggal)
    except ValueError:
        log.error("Format tanggal harus YYYY-MM-DD, dapat: CLI=%s / config TANGGAL_RAPOT (cek config.py)",
                  args.tanggal)
        return 2
    tgl_cetak = format_tanggal_id(dt)

    try:
        df, info = load_database(args.database, kelas_filter=args.kelas)
    except Exception as e:
        log.error("Gagal membaca %s: %s", args.database, e)
        return 2

    log.info("Data: %d baris (dibuang tanpa nama: %d) | titimangsa: %s (sumber: %s) | KKM: %d",
             info["jml_awal"], info["dibuang_tanpa_nama"], tgl_cetak, sumber_tgl, args.kkm)
    if args.kelas:
        log.info('Filter kelas: %s -> %d siswa', args.kelas, len(df))
    if args.limit:
        df = df.head(args.limit)

    with open(args.template, 'rb') as f:
        template_data = f.read()
    tmpl = DocxTemplate(template_data)
    try:
        tmpl.validate()
    except ValueError as e:
        log.error("%s", e)
        return 2
    log.info("Template OK: %d paragraf", tmpl.jumlah_paragraf)

    if not args.dry_run:
        os.makedirs(args.output, exist_ok=True)

    ok = gagal = dilewati = 0
    # Hanya nama yang dibuat DALAM run ini (untuk deduplikasi nama kembar).
    # File yang sudah ada di folder TIDAK dimasukkan: ia ditangani oleh
    # logika lewati/--overwrite di bawah (bukan dibuatkan _2 diam-diam).
    terpakai = set()

    for _, row in df.iterrows():
        nama_file = buat_nama_file(row.get('Kelas', ''), row.get('Nama', ''), terpakai)
        terpakai.add(nama_file)
        tujuan = os.path.join(args.output, nama_file)
        if os.path.exists(tujuan) and not args.overwrite and not args.dry_run:
            log.warning("Lewati (sudah ada, pakai --overwrite untuk tulis ulang): %s", nama_file)
            dilewati += 1
            continue
        try:
            data, info_row = tmpl.render(row, tgl_cetak, kkm=args.kkm)
            if info_row["gagal_mapel"]:
                log.warning("%s: mapel non-angka %s -> ditulis apa adanya",
                            nama_file, info_row["gagal_mapel"])
            ph = info_row.get("placeholder", {})
            if any(v == 0 for v in ph.values()):
                hilang = [k for k, v in ph.items() if v == 0]
                log.warning("%s: placeholder tidak ketemu %s (cek template)", nama_file, hilang)
            if not args.dry_run:
                with open(tujuan, 'wb') as f:
                    f.write(data)
            log.info("%s: %s (total=%s)", "Cek" if args.dry_run else "Cetak", nama_file, info_row["total"])
            ok += 1
        except Exception as e:
            log.error("Gagal %s: %s", nama_file, e)
            gagal += 1

    log.info("Selesai. berhasil=%d gagal=%d dilewati=%d%s",
             ok, gagal, dilewati, " (DRY-RUN, tidak ada file ditulis)" if args.dry_run else f" -> {args.output}")

    if args.gabung:
        if args.dry_run:
            log.info("--gabung diabaikan karena --dry-run (tidak ada file untuk digabung).")
        elif gagal and ok == 0:
            log.error("Gagal total, penggabungan dibatalkan.")
            return 1
        else:
            try:
                from gabung_rapot import gabung_rapot
                log.info("Menggabungkan per kelas: %s -> %s", args.output, args.folder_gabungan)
                hasil_gabung = gabung_rapot(args.output, args.folder_gabungan,
                                            kelas_filter=args.kelas)
                if not hasil_gabung:
                    log.warning("Penggabungan menghasilkan 0 kelas.")
                else:
                    log.info("Gabung selesai: %d kelas -> %s", len(hasil_gabung), args.folder_gabungan)
            except ImportError as e:
                log.error("Gagal gabung (dependensi kurang: %s). Jalankan: pip install -r requirements.txt", e)
                return 2
    return 1 if gagal else 0


if __name__ == "__main__":
    sys.exit(main())
