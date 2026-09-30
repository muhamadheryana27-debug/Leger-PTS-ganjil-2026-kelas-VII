"""Gabung raport individu -> satu file per kelas.

Alur resmi (dua folder terpisah):
    1. python cetak_rapot.py --tanggal ...   # -> hasil_rapot_docx/ (individu)
    2. python gabung_rapot.py                # -> hasil_rapot_gabungan/ (per kelas)
    atau sekaligus:
       python cetak_rapot.py --gabung

Folder default diambil dari config.py (OUTPUT_INDIVIDU_DIR /
OUTPUT_GABUNGAN_DIR) sehingga konsisten dengan cetak_rapot.py.
"""

import argparse
import glob
import logging
import os
import sys

from config import OUTPUT_GABUNGAN_DIR, OUTPUT_INDIVIDU_DIR

log = logging.getLogger("gabung_rapot")


def kelompokkan_per_kelas(daftar_file):
    """Kelompokkan path by prefix sebelum '_' (mis. VIIA_NAMA.docx -> VIIA)."""
    kelas_dict = {}
    for f in daftar_file:
        kelas = os.path.basename(f).split('_')[0]
        kelas_dict.setdefault(kelas, []).append(f)
    return kelas_dict


def gabung_satu_kelas(daftar_file, tujuan):
    """Gabung daftar file (urut abjad) -> satu docx. Return jumlah siswa."""
    from docx import Document
    from docxcompose.composer import Composer

    daftar_file = sorted(daftar_file)
    master = Document(daftar_file[0])
    composer = Composer(master)
    for file_siswa in daftar_file[1:]:
        doc_lanjutan = Document(file_siswa)
        master.add_page_break()
        composer.append(doc_lanjutan)
    composer.save(tujuan)
    return len(daftar_file)


def gabung_rapot(folder_input=None, folder_output=None, kelas_filter=None):
    """Fungsi utama (bisa diimport cetak_rapot.py --gabung).

    Returns:
        dict {kelas: {'siswa': int, 'file': str}} untuk kelas yang berhasil.
    """
    folder_input = folder_input or OUTPUT_INDIVIDU_DIR
    folder_output = folder_output or OUTPUT_GABUNGAN_DIR
    os.makedirs(folder_output, exist_ok=True)

    semua_file = glob.glob(os.path.join(folder_input, "*.docx"))
    if not semua_file:
        log.warning("Tidak ada file .docx di folder %s", folder_input)
        return {}

    hasil = {}
    for kelas, daftar in sorted(kelompokkan_per_kelas(semua_file).items()):
        if kelas_filter and kelas != kelas_filter_normalisasi(kelas_filter):
            continue
        tujuan = os.path.join(folder_output, f"CETAK_SEKALIGUS_{kelas}.docx")
        try:
            n = gabung_satu_kelas(daftar, tujuan)
            log.info("Kelas %s: %d siswa -> %s", kelas, n, tujuan)
            hasil[kelas] = {"siswa": n, "file": tujuan}
        except Exception as e:
            log.error("Gagal menggabung kelas %s: %s", kelas, e)
    return hasil


def kelas_filter_normalisasi(kelas):
    """'VII A' -> 'VIIA' agar cocok dengan prefix nama file."""
    return str(kelas).replace(" ", "").upper()


def build_parser():
    p = argparse.ArgumentParser(description="Gabung raport individu per kelas.")
    p.add_argument("--input", default=OUTPUT_INDIVIDU_DIR,
                   help="Folder file individu (default dari config.py)")
    p.add_argument("--output", default=OUTPUT_GABUNGAN_DIR,
                   help="Folder file gabungan (default dari config.py)")
    p.add_argument("--kelas", default=None,
                   help='Hanya gabung satu kelas, mis. "VII A" (dicocokkan ke VIIA)')
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv=None):
    logging.basicConfig(level=logging.DEBUG if (argv and ("-v" in argv or "--verbose" in argv))
                        else logging.INFO, format="%(levelname)s: %(message)s")
    args = build_parser().parse_args(argv)
    print("=== PROSES MENGGABUNGKAN RAPORT PER KELAS ===")
    print(f"Input : {args.input}\nOutput: {args.output}")
    try:
        hasil = gabung_rapot(args.input, args.output, kelas_filter=args.kelas)
    except ImportError as e:
        log.error("Dependensi belum terinstall (%s). Jalankan: pip install -r requirements.txt", e)
        return 2
    if not hasil:
        print("Tidak ada kelas yang digabung. Cek folder input / filter --kelas.")
        return 1
    print(f"\nSelesai: {len(hasil)} kelas, {sum(v['siswa'] for v in hasil.values())} siswa.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
