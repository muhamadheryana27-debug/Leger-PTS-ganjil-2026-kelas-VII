import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

from config import DAFTAR_MAPEL, DATABASE_DEFAULT, KKM

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Dashboard Raport", layout="wide")
st.title("📊 Dashboard Evaluasi PSTS & Leger Kelas")

# --- MEMBACA DAN MENGOLAH DATA ---
@st.cache_data
def load_data():
    engine = 'openpyxl' if DATABASE_DEFAULT.lower().endswith('.xlsx') else 'odf'
    df = pd.read_excel(DATABASE_DEFAULT, engine=engine, dtype={'NIP': str, 'NIS/NISN': str, 'Kelas': str})
    mapel = DAFTAR_MAPEL
    
    for m in mapel:
        df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
        
    df['Total Nilai'] = df[mapel].sum(axis=1)
    df['Rata-rata'] = df[mapel].mean(axis=1).round(2)
    df = df.dropna(subset=['Nama'])
    return df, mapel

df, daftar_mapel = load_data()

# --- SIDEBAR PENGATURAN KELAS ---
st.sidebar.header("⚙️ Pengaturan")
daftar_kelas = sorted(df['Kelas'].dropna().unique())
kelas_pilihan = st.sidebar.selectbox("Pilih Kelas:", daftar_kelas)

df_kelas = df[df['Kelas'] == kelas_pilihan].copy()
df_kelas = df_kelas.sort_values(by='Total Nilai', ascending=False).reset_index(drop=True)
df_kelas.index = df_kelas.index + 1 

# --- FUNGSI UNDUH EXCEL ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Leger')
    return output.getvalue()

# --- TABS (NAVIGASI) ---
# Membagi halaman menjadi 3 tab agar lebih rapi
tab1, tab2, tab3 = st.tabs(["🏆 Top 10", "📋 Leger Lengkap", "👤 Profil Siswa"])

# --- TAB 1: TOP 10 ---
with tab1:
    st.header(f"Top 10 Siswa - Kelas {kelas_pilihan}")
    top_10 = df_kelas.head(10)
    fig_bar = px.bar(top_10, x='Nama', y='Total Nilai', text='Total Nilai', color='Total Nilai', color_continuous_scale='Blues')
    st.plotly_chart(fig_bar, use_container_width=True)

# --- TAB 2: LEGER LENGKAP ---
with tab2:
    st.header(f"Leger Lengkap - Kelas {kelas_pilihan}")
    kolom_tampil = ['NIS/NISN', 'Nama'] + daftar_mapel + ['Total Nilai', 'Rata-rata']
    data_xlsx = to_excel(df_kelas[kolom_tampil])
    
    st.download_button(
        label="📥 Unduh Leger (Excel .xlsx)",
        data=data_xlsx,
        file_name=f"Leger_Kelas_{kelas_pilihan}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.dataframe(df_kelas[kolom_tampil], use_container_width=True)

# --- TAB 3: PROFIL SISWA (FITUR BARU) ---
with tab3:
    st.header("Detail Prestasi Siswa")
    
    # Dropdown untuk memilih siswa di kelas tersebut
    daftar_siswa = df_kelas['Nama'].tolist()
    siswa_pilihan = st.selectbox("Cari / Pilih Nama Siswa:", daftar_siswa)
    
    if siswa_pilihan:
        # Ambil data spesifik siswa tersebut
        data_siswa = df_kelas[df_kelas['Nama'] == siswa_pilihan].iloc[0]
        
        # Cari peringkat siswa di kelas
        peringkat = df_kelas[df_kelas['Nama'] == siswa_pilihan].index[0]
        
        st.markdown("---")
        
        # Membagi layar menjadi dua kolom
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Informasi Umum")
            st.write(f"**Nama:** {data_siswa['Nama']}")
            st.write(f"**NIS/NISN:** {data_siswa['NIS/NISN']}")
            st.write(f"**Kelas:** {data_siswa['Kelas']}")
            st.write(f"**Wali Kelas:** {data_siswa['Wali Kelas']}")
            
            st.markdown("<br>", unsafe_allow_html=True) # Spasi
            
            # Highlight Nilai
            st.metric(label="Peringkat di Kelas", value=f"Ke-{peringkat} dari {len(df_kelas)}")
            st.metric(label="Total Nilai", value=data_siswa['Total Nilai'])
            st.metric(label="Rata-rata", value=data_siswa['Rata-rata'])

        with col2:
            st.subheader("Peta Kekuatan Nilai (Radar Chart)")
            
            # Menyiapkan data untuk Radar Chart
            nilai_siswa = [data_siswa[m] for m in daftar_mapel]
            
            fig_radar = go.Figure()
            fig_radar.add_trace(go.Scatterpolar(
                r=nilai_siswa + [nilai_siswa[0]], # Ditutup ke titik awal
                theta=daftar_mapel + [daftar_mapel[0]],
                fill='toself',
                name=siswa_pilihan,
                line_color='#4C72B0'
            ))
            
            fig_radar.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 100]) # Skala 0-100
                ),
                showlegend=False,
                margin=dict(l=40, r=40, t=20, b=20)
            )
            st.plotly_chart(fig_radar, use_container_width=True)
            
        st.markdown("---")
        
        # Rincian Nilai dan Ketuntasan (KKM dari config.py)
        st.subheader("Rincian Nilai & Status Ketuntasan")
        
        # Buat tabel kecil untuk rincian status
        rincian_data = []
        for m in daftar_mapel:
            nilai = data_siswa[m]
            status = "✅ Tuntas" if nilai >= KKM else "❌ Belum Tuntas"
            rincian_data.append({"Mata Pelajaran": m, "Nilai": nilai, "Status": status})
            
        df_rincian = pd.DataFrame(rincian_data)
        
        # Hitung berapa mapel yang belum tuntas
        jumlah_remedial = len(df_rincian[df_rincian['Status'] == "❌ Belum Tuntas"])
        
        if jumlah_remedial > 0:
            st.warning(f"Siswa ini memiliki {jumlah_remedial} mata pelajaran di bawah KKM dan perlu bimbingan.")
        else:
            st.success("Siswa ini tuntas pada semua mata pelajaran. Pertahankan!")
            
        st.dataframe(df_rincian, use_container_width=True, hide_index=True)