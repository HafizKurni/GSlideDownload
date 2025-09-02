import os
import io
import shutil
import zipfile
import requests
import streamlit as st
from pdf2image import convert_from_bytes

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(
    page_title="Pengunduh Google Slides",
    page_icon="📄",
    layout="centered"
)

# --- UI Aplikasi Utama ---
st.title("📄 Pengunduh Google Slides")
st.write("Tempelkan URL 'embed' dari presentasi Google Slides publik untuk mengunduh semua slide sebagai file ZIP berisi gambar PNG.")

# --- Konfigurasi ---
DEFAULT_URL = "https://docs.google.com/presentation/d/e/2PACX-1vRoiaaNbJnyMe-Z19h9h2wy24yJsX_rFhHn6_svn5VMKMDcl4yMsjMF3qoNUOO_Yg/embed?start=false&loop=false&delayms=3000&slide=id.p1"
OUTPUT_FOLDER = "temp_slides"
ZIP_FILENAME = "downloaded_slides.zip"

def convert_url_to_pdf_export(url):
    """Mengubah URL Google Slides dari format /embed menjadi /export/pdf."""
    try:
        # Menemukan bagian dasar dari URL
        base_url = url.split('/embed')[0]
        return f"{base_url}/export/pdf"
    except Exception:
        return None

def download_and_convert(url):
    """Mengunduh file PDF dan mengubahnya menjadi gambar."""
    pdf_url = convert_url_to_pdf_export(url)
    if not pdf_url:
        st.error("URL yang Anda masukkan tidak valid. Pastikan ini adalah tautan 'embed' Google Slides.")
        return False

    status_placeholder = st.empty()
    progress_bar = st.progress(0, text="Mempersiapkan unduhan...")

    try:
        # Hapus folder output lama jika ada
        if os.path.exists(OUTPUT_FOLDER):
            shutil.rmtree(OUTPUT_FOLDER)
        os.makedirs(OUTPUT_FOLDER)

        # 1. Unduh PDF
        status_placeholder.info(f"Mengunduh file PDF dari Google...")
        response = requests.get(pdf_url, stream=True)
        response.raise_for_status()  # Cek jika ada error HTTP
        pdf_bytes = response.content
        
        # 2. Konversi PDF menjadi gambar
        status_placeholder.info("Mengonversi PDF menjadi gambar... Ini mungkin memakan waktu beberapa saat.")
        images = convert_from_bytes(pdf_bytes, fmt='png')
        
        total_slides = len(images)
        status_placeholder.info(f"Ditemukan {total_slides} slide. Menyimpan gambar...")

        for i, image in enumerate(images):
            file_path = os.path.join(OUTPUT_FOLDER, f"slide_{str(i + 1).zfill(3)}.png")
            image.save(file_path, "PNG")
            progress_bar.progress((i + 1) / total_slides, text=f"Menyimpan gambar {i+1}/{total_slides}")

        status_placeholder.success("Semua slide berhasil diubah menjadi gambar! Sekarang membuat file ZIP...")
        return True

    except requests.exceptions.RequestException as e:
        st.error(f"Gagal mengunduh PDF. Pastikan presentasi ini bersifat publik. Error: {e}")
        return False
    except Exception as e:
        st.error(f"Terjadi kesalahan saat konversi: {e}")
        return False

def zip_files():
    """Membuat file ZIP dari gambar yang diunduh."""
    with zipfile.ZipFile(ZIP_FILENAME, 'w') as zipf:
        for root, _, files in os.walk(OUTPUT_FOLDER):
            for file in sorted(files):
                zipf.write(os.path.join(root, file), arcname=file)

# --- UI dan Logika Streamlit ---
slides_url = st.text_input("Masukkan URL Embed Google Slides", DEFAULT_URL)

if st.button("Unduh Slide", type="primary"):
    if slides_url:
        download_success = download_and_convert(slides_url)
        
        if download_success:
            zip_files()
            with open(ZIP_FILENAME, "rb") as fp:
                st.download_button(
                    label="Unduh ZIP",
                    data=fp,
                    file_name="slides.zip",
                    mime="application/zip",
                )
            # Membersihkan file sementara setelah tombol unduh ditampilkan
            shutil.rmtree(OUTPUT_FOLDER)
            os.remove(ZIP_FILENAME)
    else:
        st.warning("Silakan masukkan URL.")

