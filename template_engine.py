"""Engine template DOCX (lxml + zipfile) untuk template placeholder baru.

Template memakai dua sistem:
1. Placeholder UNIK ([Kelas], [Nama], [NIS/NISN], [jumlah], [NAMA_WALI],
   [NIP_WALI]) -> diisi by pencarian teks, tahan geser paragraf.
2. Placeholder GENERIK berulang ([angka]/[huruf]/[deskripsi]) -> diisi by
   nomor paragraf via MAPEL_SLOT di config.py.

Bold yang dipasang di template DIPERTAHANKAN otomatis (auto-detect per
paragraf) sehingga teks pengganti tetap tebal seperti placeholder-nya.
"""

import io
import zipfile

from lxml import etree

from config import (
    KKM, MAPEL_SLOT, MARKER_ANGKA, MARKER_DESKRIPSI,
    MARKER_HURUF, NS, PH_JUMLAH, PH_KELAS, PH_NAMA, PH_NAMA_WALI,
    PH_NIP_WALI, PH_NIS, PH_TANGGAL, TOTAL_HURUF, XML_SPACE,
)
from utils import (
    angka_ke_huruf, bersihkan_nis, format_nip_display,
    nilai_ke_int, status_tuntas,
)


def _tulis_paragraf(paragraf, teks, bold=False, underline=False):
    """Tulis ulang isi satu <w:p> dengan satu <w:r> baru."""
    for r in paragraf.findall(f'{{{NS}}}r'):
        paragraf.remove(r)
    new_r = etree.SubElement(paragraf, f'{{{NS}}}r')
    if bold or underline:
        rPr = etree.SubElement(new_r, f'{{{NS}}}rPr')
        if bold:
            etree.SubElement(rPr, f'{{{NS}}}b')
        if underline:
            u = etree.SubElement(rPr, f'{{{NS}}}u')
            u.set(f'{{{NS}}}val', 'single')
    t = etree.SubElement(new_r, f'{{{NS}}}t')
    # Penting: spasi di awal (": Nama") hanya awet jika preserve
    t.set(f'{{{XML_SPACE}}}space', 'preserve')
    t.text = str(teks)


def _teks_paragraf(paragraf):
    return "".join(
        t.text for t in paragraf.findall(f'.//{{{NS}}}t') if t.text
    )


def _gaya_paragraf(paragraf):
    """Deteksi bold/underline yang dipakai paragraf (untuk dipertahankan).

    Placeholder di template dicetak tebal; tanpa ini teks pengganti
    akan tercetak tipis dan berbeda dari desain template.
    """
    bold = any(r.find(f'{{{NS}}}rPr/{{{NS}}}b') is not None
               for r in paragraf.findall(f'{{{NS}}}r'))
    ul = any(r.find(f'{{{NS}}}rPr/{{{NS}}}u') is not None
             for r in paragraf.findall(f'{{{NS}}}r'))
    return bold, ul


class DocxTemplate:
    """Template yang sudah dimuat di memori, siap render per siswa."""

    def __init__(self, template_bytes):
        self._template_bytes = bytes(template_bytes)
        with zipfile.ZipFile(io.BytesIO(self._template_bytes)) as zin:
            try:
                xml_content = zin.read('word/document.xml')
            except KeyError as e:
                raise ValueError("Template bukan DOCX valid (word/document.xml hilang)") from e
            self._items = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        self._tree = etree.fromstring(xml_content)
        self._paras = self._tree.findall(f'.//{{{NS}}}p')

    @property
    def jumlah_paragraf(self):
        return len(self._paras)

    def validate(self):
        """Pastikan template cocok dengan MAPEL_SLOT di config.py.

        Selain jumlah paragraf, tiap slot dicek MASIH berisi marker yang
        diharapkan ([angka]/[huruf]/[deskripsi]). Jika template diedit dan
        sel bergeser, error di sini — bukan rapor salah isi diam-diam.
        """
        maks = max(
            [TOTAL_HURUF]
            + [idx for slot in MAPEL_SLOT.values() for idx in slot.values()]
        )
        if len(self._paras) <= maks:
            raise ValueError(
                f"Template hanya punya {len(self._paras)} paragraf, "
                f"tapi konfigurasi butuh indeks {maks}. "
                f"Template kemungkinan berubah — cek ulang MAPEL_SLOT di config.py."
            )
        marker = {'angka': MARKER_ANGKA, 'huruf': MARKER_HURUF,
                  'deskripsi': MARKER_DESKRIPSI}
        for mapel, slot in MAPEL_SLOT.items():
            for field, idx in slot.items():
                if marker[field] not in _teks_paragraf(self._paras[idx]):
                    raise ValueError(
                        f"Slot {mapel}.{field} (paragraf {idx}) tidak berisi "
                        f"'{marker[field]}' lagi: {_teks_paragraf(self._paras[idx])[:40]!r}. "
                        f"Template berubah — perbarui MAPEL_SLOT di config.py."
                    )
        for nama, idx, mark in [("Total-huruf", TOTAL_HURUF, MARKER_HURUF)]:
            if mark not in _teks_paragraf(self._paras[idx]):
                raise ValueError(
                    f"Slot {nama} (paragraf {idx}) tidak berisi '{mark}' lagi. "
                    f"Template berubah — perbarui config.py."
                )
        return True

    # -- primitif --
    def set_para_text(self, idx, text, bold=None, underline=None):
        """Tulis paragraf by indeks. bold=None -> pertahankan gaya template.

        Return False jika indeks di luar jangkauan.
        """
        if idx >= len(self._paras):
            return False
        p = self._paras[idx]
        if bold is None or underline is None:
            b_auto, u_auto = _gaya_paragraf(p)
            bold = b_auto if bold is None else bold
            underline = u_auto if underline is None else underline
        _tulis_paragraf(p, text, bold=bold, underline=underline)
        return True

    def ganti_placeholder(self, placeholder, pengganti, bold=None, underline=None):
        """Ganti placeholder UNIK di semua paragraf (bold=None -> auto).

        Return jumlah paragraf yang kena.
        """
        kena = 0
        for p in self._paras:
            full = _teks_paragraf(p)
            if placeholder in full:
                if bold is None or underline is None:
                    b_auto, u_auto = _gaya_paragraf(p)
                    b = b_auto if bold is None else bold
                    u = u_auto if underline is None else underline
                else:
                    b, u = bold, underline
                _tulis_paragraf(p, full.replace(placeholder, str(pengganti)),
                                bold=b, underline=u)
                kena += 1
        return kena

    # -- pengisian per bagian --
    def isi_identitas(self, row):
        # Template sudah memuat ": " -> ganti placeholder-nya saja
        self.ganti_placeholder(PH_KELAS, str(row.get('Kelas', '')).strip())
        self.ganti_placeholder(PH_NAMA, str(row.get('Nama', '')).upper().strip())
        self.ganti_placeholder(PH_NIS, bersihkan_nis(row.get('NIS/NISN', '')))

    def isi_nilai(self, row, kkm=KKM):
        """Isi semua mapel + Jumlah. Return (total, gagal).

        Nilai kosong (NaN/'') -> '-' di ketiga sel (tidak dilog).
        Teks non-angka -> angka apa adanya, huruf/deskripsi '-', dicatat
        di `gagal` untuk dilog sebagai peringatan.
        """
        total = 0
        gagal = []
        for col, slot in MAPEL_SLOT.items():
            ok, v = nilai_ke_int(row.get(col, None))
            if ok:
                self.set_para_text(slot['angka'], str(v))
                self.set_para_text(slot['huruf'], angka_ke_huruf(v))
                self.set_para_text(slot['deskripsi'], status_tuntas(v, kkm))
                total += v
            else:
                mentah = row.get(col, '')
                if mentah is None or (isinstance(mentah, float) and mentah != mentah):
                    mentah = ''
                if str(mentah).strip() in ('', '0'):
                    self.set_para_text(slot['angka'], '-')
                    self.set_para_text(slot['huruf'], '-')
                    self.set_para_text(slot['deskripsi'], '-')
                else:
                    self.set_para_text(slot['angka'], str(mentah))
                    self.set_para_text(slot['huruf'], '-')
                    self.set_para_text(slot['deskripsi'], '-')
                    gagal.append(col)
        # Keterampilan & BTQ di template sudah '-' (literal) -> tak disentuh.
        # Jumlah angka + terbilang
        self.ganti_placeholder(PH_JUMLAH, str(total))
        self.set_para_text(TOTAL_HURUF, angka_ke_huruf(total))
        return total, gagal

    def isi_wali_kelas(self, row, tgl_cetak):
        nama_w = str(row.get('Wali Kelas', '')).strip()
        nip_str = format_nip_display(row.get('NIP', ''))
        # Tanpa bold/underline eksplisit: gaya mengikuti template
        # (saat ini [NAMA_WALI]/[NIP_WALI] tebal+garis bawah di template).
        c1 = self.ganti_placeholder(PH_TANGGAL, tgl_cetak)
        if c1 == 0:
            # Template tidak punya [TITI_MANGSA]; tanggal masih hardcoded
            # ("Wanayasa, ...") -> timpa paragraf titimangsa yang ada.
            c1 = self._tulis_titimangsa_fallback(tgl_cetak)
        c2 = self.ganti_placeholder(PH_NAMA_WALI, nama_w)
        c3 = self.ganti_placeholder(PH_NIP_WALI, nip_str)
        return {PH_TANGGAL: c1, PH_NAMA_WALI: c2, PH_NIP_WALI: c3}

    def _tulis_titimangsa_fallback(self, tgl_cetak):
        """Timpa paragraf pertama yang diawali 'Wanayasa,'. Return 1/0.

        Gaya (tebal/garis bawah) mengikuti template, bukan hardcoded.
        """
        for p in self._paras:
            full = _teks_paragraf(p).strip()
            if full.startswith("Wanayasa,"):
                b, u = _gaya_paragraf(p)
                _tulis_paragraf(p, tgl_cetak, bold=b, underline=u)
                return 1
        return 0

    def render(self, row, tgl_cetak, kkm=KKM):
        """Isi 1 siswa -> kembalikan (docx_bytes, info_dict)."""
        # Render harus idempoten per siswa: segarkan tree dari template asli
        # agar objek DocxTemplate bisa dipakai ulang untuk banyak siswa.
        # (Refresh DULU, baru validate — validate membaca tree aktif.)
        self._tree = etree.fromstring(self._items['word/document.xml'])
        self._paras = self._tree.findall(f'.//{{{NS}}}p')
        self.validate()
        self.isi_identitas(row)
        total, gagal = self.isi_nilai(row, kkm=kkm)
        ph = self.isi_wali_kelas(row, tgl_cetak)
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for name, data in self._items.items():
                if name == 'word/document.xml':
                    zout.writestr(name, etree.tostring(
                        self._tree, xml_declaration=True,
                        encoding='UTF-8', standalone=True))
                else:
                    zout.writestr(name, data)
        return out.getvalue(), {"total": total, "gagal_mapel": gagal, "placeholder": ph}


# -- Wrapper kompatibel dengan API lama --
def buat_file_rapot(template_bytes, row, tgl_cetak=None, kkm=KKM):
    """API lama: buat_file_rapot(template_bytes, row) -> bytes.

    tgl_cetak opsional agar cetak ulang deterministik (dulu selalu now()).
    """
    from utils import format_tanggal_id
    tmpl = DocxTemplate(template_bytes)
    data, _info = tmpl.render(row, tgl_cetak or format_tanggal_id(), kkm=kkm)
    return data
