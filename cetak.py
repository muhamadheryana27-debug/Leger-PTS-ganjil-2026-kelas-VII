import os
import datetime
import pandas as pd
from lxml import etree
import zipfile
import io

# ── KONFIGURASI ──────────────────────────────────────────────────────────────
TEMPLATE_DOCX = "Print_Rapot.docx"
DATABASE_ODS  = "database.ods"
OUTPUT_DIR    = "hasil_rapot_docx"

BULAN_ID = {
    1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni",
    7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"
}

# Namespace OpenXML 
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = 'http://www.w3.org/XML/1998/namespace'

# ── FUNGSI PENDUKUNG ─────────────────────────────────────────────────────────

def angka_ke_huruf(n):
    """Fungsi Terbilang Bahasa Indonesia"""
    satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas']
    try:
        n = int(round(float(n)))
        if n < 0: return '-'
        if n < 12: return satuan[n]
        if n < 20: return satuan[n-10] + ' Belas'
        if n < 100: return (satuan[n//10] + ' Puluh ' + satuan[n%10]).strip()
        if n < 200: return ('Seratus ' + angka_ke_huruf(n-100)).strip()
        if n < 1000: return (satuan[n//100] + ' Ratus ' + angka_ke_huruf(n%100)).strip()
        return str(n)
    except: return '-'

def set_identity_run(p, value, bold_value=False):
    """
    Mengatur paragraf identitas secara presisi.
    Mereset isi paragraf agar format ': ' (titik dua) normal, 
    dan variabel value bisa normal/tebal (menjaga agar titik dua lurus sejajar).
    """
    for r in p.findall(f'.//{{{NS}}}r'): 
        p.remove(r)
    
    # Run 1: Titik dua (Tidak ditebalkan)
    r1 = etree.SubElement(p, f'{{{NS}}}r')
    t1 = etree.SubElement(r1, f'{{{NS}}}t')
    t1.set(f'{{{XML_SPACE}}}space', 'preserve')
    t1.text = ": "
    
    # Run 2: Isi Identitas (Opsi Ditebalkan)
    r2 = etree.SubElement(p, f'{{{NS}}}r')
    if bold_value:
        rPr = etree.SubElement(r2, f'{{{NS}}}rPr')
        etree.SubElement(rPr, f'{{{NS}}}b')
    t2 = etree.SubElement(r2, f'{{{NS}}}t')
    t2.set(f'{{{XML_SPACE}}}space', 'preserve')
    t2.text = str(value)

def replace_full_paragraph(p, text, bold=False, underline=False):
    """Mengganti seluruh isi paragraf placeholder murni menjadi teks nilai/data."""
    for r in p.findall(f'.//{{{NS}}}r'): 
        p.remove(r)
        
    new_r = etree.SubElement(p, f'{{{NS}}}r')
    if bold or underline:
        rPr = etree.SubElement(new_r, f'{{{NS}}}rPr')
        if bold: etree.SubElement(rPr, f'{{{NS}}}b')
        if underline:
            u = etree.SubElement(rPr, f'{{{NS}}}u')
            u.set(f'{{{NS}}}val', 'single')
            
    new_t = etree.SubElement(new_r, f'{{{NS}}}t')
    new_t.set(f'{{{XML_SPACE}}}space', 'preserve')
    new_t.text = str(text)

# ── PROSES INTI ──────────────────────────────────────────────────────────────

def buat_file_rapot(template_bytes, row):
    now = datetime.datetime.now()
    tgl_cetak = f"Wanayasa, {now.day} {BULAN_ID[now.month]} {now.year}"
    
    # -- Sanitasi Data Database --
    kelas = str(row.get('Kelas', ''))
    nama  = str(row.get('Nama', '')).upper()
    
    nis_raw = str(row.get('NIS/NISN', '')).strip()
    if nis_raw.endswith('.0'): nis_raw = nis_raw[:-2]
    nis = "-" if nis_raw.lower() in ['nan', '', 'none'] else nis_raw
    
    nama_wali = str(row.get('Wali Kelas', '')).strip()
    nip_raw   = str(row.get('NIP', '')).strip()
    if nip_raw.endswith('.0'): nip_raw = nip_raw[:-2]
    nip_wali  = f"NIP. {nip_raw}" if nip_raw.lower() not in ['nan', '', 'none'] else "NIP. -"

    # Hitung Total Nilai Dinamis
    mapel_cols = ['PAI', 'PPKn', 'B.IND', 'B.ING', 'MTK', 'IPA', 'IPS', 'SBY', 'PJOK', 'INF', 'B.SUN']
    total_nilai = 0
    for col in mapel_cols:
        val = row.get(col, 0)
        if pd.notna(val) and str(val).strip() != '':
            try: total_nilai += int(round(float(val)))
            except: pass

    # Buka XML Document
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        xml_content = zin.read('word/document.xml')
    
    tree = etree.fromstring(xml_content)
    paras = tree.findall(f'.//{{{NS}}}p')

    # State Machine untuk mendeteksi sedang berada di baris mapel apa
    current_mapel = None

    # Iterasi dari atas ke bawah dokumen
    for p in paras:
        texts = [t.text for t in p.findall(f'.//{{{NS}}}t') if t.text]
        full_text = "".join(texts)
        if not full_text: continue

        # 1. REPLACE IDENTITAS SISWA (Menggunakan fungsi presisi)
        if "[Kelas]" in full_text:
            set_identity_run(p, kelas, bold_value=False)
            continue
        elif "[Nama]" in full_text:
            set_identity_run(p, nama, bold_value=True)
            continue
        elif "[NIS/NISN]" in full_text:
            set_identity_run(p, nis, bold_value=False)
            continue
        
        # 2. DETEKSI MATA PELAJARAN (State Machine)
        if "Agama Islam" in full_text: current_mapel = 'PAI'
        elif "Pancasila" in full_text: current_mapel = 'PPKn'
        elif "Bahasa Indonesia" in full_text: current_mapel = 'B.IND'
        elif "Bahasa Inggris" in full_text: current_mapel = 'B.ING'
        elif "Matematika" in full_text: current_mapel = 'MTK'
        elif "Pengetahuan Alam" in full_text: current_mapel = 'IPA'
        elif "Pengetahuan Sosial" in full_text: current_mapel = 'IPS'
        elif "Seni Budaya" in full_text: current_mapel = 'SBY'
        elif "Pendidikan Jasmani" in full_text: current_mapel = 'PJOK'
        elif "Informatika" in full_text: current_mapel = 'INF'
        elif "Bahasa Sunda" in full_text: current_mapel = 'B.SUN'

        # 3. REPLACE NILAI BERDASARKAN MAPEL TERDETEKSI
        if current_mapel and ("[angka]" in full_text or "[huruf]" in full_text or "[deskripsi]" in full_text):
            val = row.get(current_mapel, '')
            
            # Jika sel di database ODS kosong, kosongkan sel di rapot
            if pd.isna(val) or str(val).strip() == '':
                v_str, huruf, desk = "", "", ""
            else:
                try:
                    v = int(round(float(val)))
                    v_str = str(v)
                    huruf = angka_ke_huruf(v)
                    desk  = "Tuntas" if v >= 71 else "Belum Tuntas"
                except:
                    v_str, huruf, desk = str(val), "-", "-"

            # Eksekusi penggantian teks di tabel
            if "[angka]" in full_text: replace_full_paragraph(p, v_str)
            elif "[huruf]" in full_text: replace_full_paragraph(p, huruf)
            elif "[deskripsi]" in full_text: replace_full_paragraph(p, desk)
            continue
            
        # 4. REPLACE KOMPONEN LAIN-LAIN
        if "[jumlah]" in full_text:
            replace_full_paragraph(p, str(total_nilai), bold=True)
        elif "[TITI_MANGSA]" in full_text:
            replace_full_paragraph(p, tgl_cetak)
        elif "[NAMA_WALI]" in full_text:
            replace_full_paragraph(p, nama_wali, bold=True, underline=True)
        elif "[NIP_WALI]" in full_text:
            replace_full_paragraph(p, nip_wali)

    # Pack XML kembali ke bentuk Word (.docx)
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        with zipfile.ZipFile(output, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, etree.tostring(tree, xml_declaration=True, encoding='UTF-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
    return output.getvalue()

def main():
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    print("Memulai proses pembuatan raport DOCX...")
    
    try:
        df = pd.read_excel(DATABASE_ODS, engine='odf', dtype={'NIP': str, 'NIS/NISN': str, 'Kelas': str})
    except Exception as e:
        print(f"Gagal membaca file {DATABASE_ODS}: {e}")
        return

    with open(TEMPLATE_DOCX, 'rb') as f:
        template_data = f.read()

    for i, row in df.iterrows():
        if pd.isna(row['Nama']): continue
        
        # Penamaan file output
        nama_bersih = str(row['Nama']).upper().replace("/", "-").strip()
        kelas_bersih = str(row['Kelas']).replace(' ','').strip()
        nama_file = f"{kelas_bersih}_{nama_bersih}.docx"
        
        print(f"Mencetak: {nama_file}")
        
        try:
            hasil = buat_file_rapot(template_data, row)
            with open(os.path.join(OUTPUT_DIR, nama_file), 'wb') as f:
                f.write(hasil)
        except Exception as e:
            print(f"Gagal mencetak rapot {nama_file}: {e}")

    print(f"\nSelesai! Seluruh rapot berhasil dicetak di folder: {OUTPUT_DIR}")

if __name__ == "__main__":
    main()