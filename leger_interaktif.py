import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os

from config import DAFTAR_MAPEL, DATABASE_DEFAULT, KKM

# ── KONFIGURASI HALAMAN STREAMLIT ───────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Leger & Evaluasi Raport",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Evaluasi PSTS & Leger Kelas")

# ── FUNGSI MEMBACA DAN MENGOLAH DATA ─────────────────────────────────────────
@st.cache_data
def load_data():
    if not os.path.exists(DATABASE_DEFAULT):
        st.error(f"File database `{DATABASE_DEFAULT}` tidak ditemukan di server!")
        st.stop()
        
    engine = 'openpyxl' if DATABASE_DEFAULT.lower().endswith('.xlsx') else 'odf'
    df = pd.read_excel(DATABASE_DEFAULT, engine=engine, dtype={'NIP': str, 'NIS/NISN': str, 'Kelas': str})
    
    mapel = DAFTAR_MAPEL
    
    # Pengecekan aman untuk mencegah KeyError pada Streamlit Cloud
    for m in mapel:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
        else:
            df[m] = 0.0
            
    df['Total Nilai'] = df[mapel].sum(axis=1)
    df['Rata-rata'] = df[mapel].mean(axis=1).round(2)
    df = df.dropna(subset=['Nama'])
    
    df['Nama'] = df['Nama'].astype(str).str.strip()
    df['Kelas'] = df['Kelas'].astype(str).str.strip()
    
    return df, mapel

try:
    df, daftar_mapel = load_data()
except Exception as e:
    st.error(f"Gagal membaca data dari database: {e}")
    st.stop()

# ── SIDEBAR PENGATURAN & FILTER KELAS ────────────────────────────────────────
st.sidebar.header("⚙️ Pengaturan & Filter")

daftar_kelas = sorted(df['Kelas'].unique())
if not daftar_kelas:
    st.warning("Tidak ada data kelas yang ditemukan di database.")
    st.stop()

kelas_pilihan = st.sidebar.selectbox("Pilih Kelas:", daftar_kelas)

df_kelas = df[df['Kelas'] == kelas_pilihan].copy()
df_kelas = df_kelas.sort_values(by='Total Nilai', ascending=False).reset_index(drop=True)
df_kelas.index = df_kelas.index + 1 

st.sidebar.markdown("---")
st.sidebar.info(f"**Total Siswa Kelas {kelas_pilihan}:** {len(df_kelas)} Siswa\n\n**KKM Sekolah:** {KKM}")

# ── FUNGSI UNDUH EXCEL ───────────────────────────────────────────────────────
def to_excel(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name=f'Leger_{kelas_pilihan}')
    return output.getvalue()

# ── TABS NAVIGASI ────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🏆 Top 10", "📋 Leger Lengkap", "👤 Profil Siswa"])

# ── TAB 1: TOP 10 ────────────────────────────────────────────────────────────
with tab1:
    st.header(f"🏆 Top 10 Siswa Berprestasi - Kelas {kelas_pilihan}")
    top_10 = df_kelas.head(10)
    
    fig_bar = px.bar(
        top_10,
        x='Nama',
        y='Total Nilai',
        text='Total Nilai',
        color='Total Nilai',
        color_continuous_scale='Blues',
        title=f"10 Nilai Total Tertinggi di Kelas {kelas_pilihan}"
    )
    fig_bar.update_traces(textposition='outside')
    fig_bar.update_layout(xaxis_title="Nama Siswa", yaxis_title="Total Nilai", showlegend=False)
    st.plotly_chart(fig_bar, use_container_width=True)

# ── TAB 2: LEGER LENGKAP ─────────────────────────────────────────────────────
with tab2:
    st.header(f"📋 Leger Nilai Lengkap - Kelas {kelas_pilihan}")
    
    kolom_tampil = ['NIS/NISN', 'Nama'] + daftar_mapel + ['Total Nilai', 'Rata-rata']
    kolom_tampil = [k for k in kolom_tampil if k in df_kelas.columns]
    
    data_xlsx = to_excel(df_kelas[kolom_tampil])
    
    st.download_button(
        label="📥 Unduh Leger (Format Excel .xlsx)",
        data=data_xlsx,
        file_name=f"Leger_Kelas_{kelas_pilihan}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.dataframe(df_kelas[kolom_tampil], use_container_width=True)

# ── TAB 3: PROFIL SISWA ──────────────────────────────────────────────────────
with tab3:
    st.header("👤 Detail Prestasi & Evaluasi Siswa")
    
    daftar_siswa = df_kelas['Nama'].tolist()
    siswa_pilihan = st.selectbox("Cari / Pilih Nama Siswa:", daftar_siswa)
    
    if siswa_pilihan:
        data_siswa = df_kelas[df_kelas['Nama'] == siswa_pilihan].iloc[0]
        peringkat = df_kelas[df_kelas['Nama'] == siswa_pilihan].index[0]
        
        st.markdown("---")
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Informasi Umum")
            st.write(f"**Nama:** {data_siswa['Nama']}")
            st.write(f"**NIS/NISN:** {data_siswa.get('NIS/NISN', '-')}")
            st.write(f"**Kelas:** {data_siswa['Kelas']}")
            st.write(f"**Wali Kelas:** {data_siswa.get('Wali Kelas', '-')}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.metric(label="Peringkat di Kelas", value=f"Ke-{peringkat} dari {len(df_kelas)}")
            st.metric(label="Total Nilai", value=data_siswa['Total Nilai'])
            st.metric(label="Rata-rata Nilai", value=data_siswa['Rata-rata'])

        with col2:
            st.subheader("Peta Kekuatan Nilai (Radar Chart)")
            nilai_siswa = [data_siswa[m] for m in daftar_mapel]
            
            fig_radar = go.Figure()
            fig_radar.add_trace(go.Scatterpolar(
                r=nilai_siswa + [nilai_siswa[0]],
                theta=daftar_mapel + [daftar_mapel[0]],
                fill='toself',
                name=siswa_pilihan,
                line_color='#4C72B0'
            ))
            
            fig_radar.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 100])
                ),
                showlegend=False,
                margin=dict(l=40, r=40, t=20, b=20)
            )
            st.plotly_chart(fig_radar, use_container_width=True)
            
        st.markdown("---")
        st.subheader(f"Rincian Nilai & Ketuntasan (KKM = {KKM})")
        
        rincian_data = []
        for m in daftar_mapel:
            nilai = data_siswa[m]
            status = "✅ Tuntas" if nilai >= KKM else "❌ Belum Tuntas"
            rincian_data.append({"Mata Pelajaran": m, "Nilai": nilai, "Status": status})
            
        df_rincian = pd.DataFrame(rincian_data)
        jumlah_remedial = len(df_rincian[df_rincian['Status'] == "❌ Belum Tuntas"])
        
        if jumlah_remedial > 0:
            st.warning(f"⚠️ Siswa ini memiliki **{jumlah_remedial}** mata pelajaran di bawah KKM dan perlu bimbingan/remedial.")
        else:
            st.success("🎉 Selamat! Siswa ini tuntas pada semua mata pelajaran.")
            
        st.dataframe(df_rincian, use_container_width=True, hide_index=True)
