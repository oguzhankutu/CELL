import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. KURUMSAL ARAYÜZ AYARLARI
# ==========================================
st.set_page_config(page_title="CELL | Satış Analiz Otomasyonu", page_icon="⚡", layout="wide")

st.title("⚡ CELL: Kurumsal Satış Analiz Sistemi")
st.markdown("ERP verilerinizi yükleyin, sütunları eşleştirin ve saniyeler içinde dinamik özet raporlarınızı oluşturun.")

# ==========================================
# HIZLANDIRICI: ÖNBELLEK (CACHE) SİSTEMİ
# ==========================================
@st.cache_data(show_spinner="Veri işleniyor (Bu işlem dosya başına bir kez yapılır)...")
def veriyi_yukle(dosya, baslik):
    _df = pd.read_excel(dosya, header=baslik)
    return _df.dropna(how='all').reset_index(drop=True)

# ==========================================
# 2. DOSYA YÜKLEME VE BAŞLIK AYARI
# ==========================================
with st.sidebar:
    st.header("⚙️ Veri Aktarımı")
    yuklenen_dosya = st.file_uploader("Excel Dosyanızı Yükleyin (.xls, .xlsx)", type=["xls", "xlsx"])
    
    st.markdown("---")
    st.write("**Tablo Başlığı Konumu**")
    st.info("Rapor formatınıza göre gerçek sütun isimlerinin bulunduğu satırı belirleyin.")
    baslik_satiri = st.number_input("Başlık Satırı", min_value=1, value=1) - 1

if yuklenen_dosya is not None:
    try:
        df_orijinal = veriyi_yukle(yuklenen_dosya, baslik_satiri)
        df = df_orijinal.copy()
        sutunlar = df.columns.tolist()

        # ==========================================
        # 3. DİNAMİK SÜTUN EŞLEŞTİRME
        # ==========================================
        st.subheader("1️⃣ Sütun Konfigürasyonu")
        st.write("Analiz edilecek veri setindeki anahtar sütunları eşleştirin.")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            temsilci_sutunu = st.selectbox("Satış Temsilcisi:", sutunlar)
            tarih_sutunu = st.selectbox("İşlem Tarihi:", sutunlar)
        with col2:
            urun_sutunu = st.selectbox("Ürün / Kategori:", sutunlar)
            tutar_sutunu = st.selectbox("Net Tutar:", sutunlar)
        with col3:
            miktar_sutunu = st.selectbox("Miktar (Birim):", sutunlar)

        # ==========================================
        # 4. FİLTRELEME
        # ==========================================
        st.subheader("2️⃣ Personel Filtresi")
        benzersiz_temsilciler = df[temsilci_sutunu].dropna().astype(str).unique().tolist()
        
        secilen_temsilciler = st.multiselect(
            "Analize dahil edilecek personelleri seçin (Boş bırakılması durumunda tüm veriler işlenir):", 
            benzersiz_temsilciler, 
            default=benzersiz_temsilciler
        )

        # ==========================================
        # 5. RAPOR ÜRETİMİ
        # ==========================================
        if st.button("🚀 Analizi Başlat ve Raporu İndir", use_container_width=True):
            with st.spinner("Matris oluşturuluyor..."):
                
                if secilen_temsilciler:
                    df_rapor = df[df[temsilci_sutunu].astype(str).isin(secilen_temsilciler)].copy()
                else:
                    df_rapor = df.copy()

                df_rapor[tutar_sutunu] = pd.to_numeric(df_rapor[tutar_sutunu], errors='coerce').fillna(0)
                df_rapor[miktar_sutunu] = pd.to_numeric(df_rapor[miktar_sutunu], errors='coerce').fillna(0)

                df_rapor[tarih_sutunu] = pd.to_datetime(df_rapor[tarih_sutunu], errors='coerce', dayfirst=True)
                ay_haritasi = {1: '1. Ocak', 2: '2. Şubat', 3: '3. Mart', 4: '4. Nisan', 5: '5. Mayıs', 6: '6. Haziran',
                               7: '7. Temmuz', 8: '8. Ağustos', 9: '9. Eylül', 10: '10. Ekim', 11: '11. Kasım', 12: '12. Aralık'}
                df_rapor['Rapor_Ayi'] = df_rapor[tarih_sutunu].dt.month.map(ay_haritasi).fillna('Tarihsiz')

                pivot_tablo = pd.pivot_table(
                    df_rapor, 
                    index=[temsilci_sutunu, urun_sutunu], 
                    columns='Rapor_Ayi', 
                    values=[miktar_sutunu, tutar_sutunu], 
                    aggfunc='sum', 
                    fill_value=0
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pivot_tablo.to_excel(writer, sheet_name='Analiz_Raporu', merge_cells=True)
                
                excel_verisi = output.getvalue()

                st.success("✅ Veri analizi başarıyla tamamlandı.")
                
                col_btn, col_empty = st.columns([1, 2])
                with col_btn:
                    st.download_button(
                        label="📥 Excel Formatında İndir",
                        data=excel_verisi,
                        file_name="CELL_Satis_Analizi.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.markdown("### 📊 Veri Önizlemesi (İlk 15 Satır)")
                st.dataframe(pivot_tablo.head(15))

    except Exception as e:
        st.error(f"Sistem Hatası: Sütun eşleştirme işlemi gerçekleştirilemedi. Başlık satırının doğru yapılandırıldığından emin olun. Hata detayı: {e}")
else:
    st.info("👈 İşleme başlamak için sol menüden Excel veri dosyanızı sisteme yükleyin.")
