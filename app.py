import streamlit as st
import pandas as pd
import zipfile
import io
import re  # Dosya isimlerini güvenli hale getirmek için

# 🎨 Streamlit Arayüzü Başlat
st.set_page_config(page_title="Excel Sheet Ayrıştırıcı", page_icon="📂", layout="centered")

st.title("📂 Excel Sheet Ayrıştırıcı 🚀")
st.write("Yüklediğiniz Excel dosyasının her sayfasını ayrı bir dosya olarak kaydedin!")

# 📤 Kullanıcıdan Excel Dosyası Alma
uploaded_file = st.file_uploader("Lütfen Excel dosyanızı yükleyin (.xlsx)", type=["xlsx"])

if uploaded_file:
    # 📂 Excel Dosyasını Okuma
    xls = pd.ExcelFile(uploaded_file)
    
    # Kullanıcıya Sheet'leri Göster
    st.write(f"**{len(xls.sheet_names)} adet sheet bulundu:** {xls.sheet_names}")

    # Kullanıcıya ZIP mi yoksa tek tek mi indirmek istediğini sor
    download_option = st.radio("İndirme Seçeneği Seçin:", ["Tek Tek", "ZIP Olarak"])

    # ZIP Dosyası İçin Bellek Alanı Aç
    zip_buffer = io.BytesIO()

    # ZIP Dosyasını oluştur
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            
            # Dosya ismini temizleyerek güvenli hale getir
            safe_sheet_name = re.sub(r'[\\/*?:"<>|]', '', sheet_name)  # Yasaklı karakterleri temizle

            # ZIP için dosyayı belleğe kaydet
            zip_file.writestr(f"{safe_sheet_name}.xlsx", output.getvalue())

            # Tek tek indirme seçeneği varsa indirme butonu göster
            if download_option == "Tek Tek":
                st.download_button(
                    label=f"📥 {safe_sheet_name}.xlsx İndir",
                    data=output,
                    file_name=f"{safe_sheet_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # ZIP Olarak İndirme Seçeneği
    if download_option == "ZIP Olarak":
        zip_buffer.seek(0)
        st.download_button(
            label="📥 Tüm Sheet'leri ZIP Olarak İndir",
            data=zip_buffer,
            file_name="excel_sheets.zip",
            mime="application/zip"
        )
