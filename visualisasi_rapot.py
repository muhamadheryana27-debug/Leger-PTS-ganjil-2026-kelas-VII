import pandas as pd
import matplotlib.pyplot as plt

from config import DAFTAR_MAPEL, DATABASE_DEFAULT

# 1. BACA DATA
file_db = DATABASE_DEFAULT
print(f"Membaca data dari {file_db}...")
engine = 'openpyxl' if file_db.lower().endswith('.xlsx') else 'odf'
df = pd.read_excel(file_db, engine=engine)

# 2. DAFTAR MATA PELAJARAN (sumber tunggal: config.py)
mapel = DAFTAR_MAPEL

# Pastikan data di kolom mapel tersebut dibaca sebagai angka
df[mapel] = df[mapel].apply(pd.to_numeric, errors='coerce')

# 3. HITUNG RATA-RATA KELAS UNTUK SETIAP MAPEL
rata_rata = df[mapel].mean().round(1)

# 4. BUAT GRAFIK (BAR CHART)
plt.figure(figsize=(12, 6)) # Mengatur ukuran gambar
bars = plt.bar(rata_rata.index, rata_rata.values, color='#4C72B0', edgecolor='black')

# Desain Tampilan Grafik
plt.title('Rata-Rata Nilai Raport Kelas per Mata Pelajaran', fontsize=16, fontweight='bold', pad=15)
plt.xlabel('Mata Pelajaran', fontsize=12, fontweight='bold')
plt.ylabel('Nilai Rata-Rata', fontsize=12, fontweight='bold')
plt.ylim(0, 100) # Membatasi sumbu Y dari 0 sampai 100
plt.grid(axis='y', linestyle='--', alpha=0.7) # Garis bantu

# Menambahkan angka teks tepat di atas setiap batang grafik
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval + 1, f'{yval}', ha='center', va='bottom', fontsize=11, fontweight='bold')

# Tampilkan grafik ke layar
plt.tight_layout()
plt.show()

# Opsional: Jika ingin langsung menyimpannya jadi gambar, gunakan kode di bawah ini dan hapus plt.show()
# plt.savefig('Grafik_Rata_Rata_Mapel.png', dpi=300)
# print("Grafik berhasil disimpan sebagai 'Grafik_Rata_Rata_Mapel.png'")