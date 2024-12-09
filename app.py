import streamlit as st
import pandas as pd
import zipfile
import os
import numpy as np
from pdf2image import convert_from_path
import cv2

# Fungsi untuk memvalidasi tanda tangan
def validate_signature(pdf_path):
    images = convert_from_path(pdf_path)
    for image in images:
        # Konversi gambar ke format yang dapat diproses OpenCV
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        # Deteksi tanda tangan (misalnya, dengan thresholding)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
        # Hitung jumlah piksel putih
        white_pixels = cv2.countNonZero(thresh)
        if white_pixels < 1000:  # Threshold untuk mendeteksi tanda tangan
            return False
    return True

# Fungsi untuk memproses file Excel
def process_excel(file):
    df = pd.read_excel(file)
    return df

# Fungsi untuk memvalidasi dokumen
def validate_documents(df, zip_file):
    feedback = []
    status = []
    
    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall("temp_folder")
        files = os.listdir("temp_folder")
        
        for index, row in df.iterrows():
            kode_mahasiswa = row['Kode Mahasiswa']
            nama_mahasiswa = row['Nama Mahasiswa']
            kode_dosen_pembimbing = row['Kode Dosen Pembimbing']
            kode_dosen_reviewer = row['Kode Dosen Reviewer']
            
            # Validasi dokumen
            documents = {
                'Proposal Dosen Pembimbing': f"{kode_mahasiswa}_{kode_dosen_pembimbing}_Dosen Pembimbing.docx",
                'Proposal Dosen Reviewer': f"{kode_mahasiswa}_{kode_dosen_reviewer}_Dosen Reviewer.docx",
                'Logbook': f"{kode_mahasiswa}_{nama_mahasiswa}_Lembar Pemantauan Bimbingan.pdf",
                'Rencana Kerja': f"{kode_mahasiswa}_{nama_mahasiswa}_Rencana Kerja Penulisan Skripsi.pdf"
            }
            
            missing_docs = []
            for doc_type, expected_name in documents.items():
                if expected_name not in files:
                    missing_docs.append(doc_type)
                else:
                    # Validasi tanda tangan untuk PDF
                    if doc_type in ['Logbook', 'Rencana Kerja']:
                        if not validate_signature(os.path.join("temp_folder", expected_name)):
                            feedback.append(f"File '{expected_name}' belum memiliki tanda tangan. Mohon pastikan semua dokumen ditandatangani sebelum diunggah.")
            
            if missing_docs:
                feedback.append(f"Dokumen berikut belum dikumpulkan oleh mahasiswa {nama_mahasiswa}: {', '.join(missing_docs)}.")
            
            status.append("Lengkap" if not missing_docs else "Tidak Lengkap")
    
    return status, feedback

# Streamlit UI
st.title("e-RP Assistant System")

# Upload file Excel
excel_file = st.file_uploader("Unggah File Excel (Master Data Mahasiswa)", type=["xlsx"])
if excel_file:
    df = process_excel(excel_file)
    st.success("File Excel berhasil diunggah.")
else:
    st.warning("File Excel tidak valid atau tidak diunggah. Mohon unggah file referensi dengan format yang benar. Unggah file dengan format Excel (.xlsx). ⚠️ File Excel harus memiliki kolom 'NamaMahasiswa', 'KodeMahasiswa', 'KodeDosenPembimbing', dan 'KodeDosenReviewer'.")

# Upload folder dokumen
zip_file = st.file_uploader("Unggah Folder Dokumen (ZIP)", type=["zip"])
if zip_file:
    with open("temp.zip", "wb") as f:
        f.write(zip_file.getbuffer())
    
    status, feedback = validate_documents(df, "temp.zip")
    
    # Tampilkan hasil
    result_df = df.copy()
    result_df['Status Dokumen'] = status
    result_df['Feedback'] = feedback
    
    st.dataframe(result_df.style.apply(lambda x: ['background: yellow' if v == 'Tidak Lengkap' else '' for v in x['Status Dokumen']], axis=1))
    
    # Tombol untuk download summary
    if st.button("Download Summary"):
        result_df.to_excel("summary.xlsx", index=False)
        st.success("Summary berhasil diunduh.")
else:
    st.warning("File dokumen tidak valid atau tidak diunggah. Mohon unggah folder dokumen dalam format ZIP.")

# ```python
# Membersihkan folder sementara setelah proses
if os.path.exists("temp_folder"):
    for file in os.listdir("temp_folder"):
        os.remove(os.path.join("temp_folder", file))
    os.rmdir("temp_folder")

# Menambahkan fitur untuk mengunduh file summary
if st.button("Download Summary"):
    result_df.to_excel("summary.xlsx", index=False)
    with open("summary.xlsx", "rb") as f:
        st.download_button("Download Summary File", f, file_name="summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
