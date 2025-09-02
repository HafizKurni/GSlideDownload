import streamlit as st
import os
import time
import zipfile
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Menggunakan cache Streamlit untuk mencegah inisialisasi ulang driver pada setiap interaksi
@st.cache_resource
def setup_driver():
    """
    Menyiapkan driver Chrome dengan opsi yang diperlukan untuk berjalan
    di lingkungan cloud Streamlit yang headless (tanpa GUI).
    """
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("window-size=1920,1080")
    
    # Menggunakan path standar di mana Streamlit Cloud menginstal chromedriver
    # Ini adalah kunci agar bisa berjalan saat di-deploy
    service = Service('/usr/bin/chromedriver') 
    
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def download_and_zip_slides(url, progress_bar, status_text):
    """
    Fungsi utama untuk mengunduh slide dengan Selenium dan mengemasnya ke dalam file ZIP.
    Bekerja dengan data di memori untuk efisiensi, tanpa menyimpan file sementara.
    """
    driver = setup_driver()
    status_text.text("Membuka URL presentasi...")
    driver.get(url)

    image_data_list = []
    
    try:
        # Menunggu elemen-elemen kunci dimuat di halaman
        wait = WebDriverWait(driver, 30)
        next_button_selector = ".punch-viewer-navbar-next"
        paginator_selector = ".punch-viewer-paginator-page-count"
        slide_container_selector = ".punch-viewer-content"

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, next_button_selector)))
        paginator_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, paginator_selector)))
        slide_container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, slide_container_selector)))
        
        # Dapatkan jumlah total slide dari teks paginator
        total_slides = int(paginator_element.text.strip())
        status_text.text(f"Ditemukan {total_slides} slide. Memulai proses...")

        # Loop untuk setiap slide
        for i in range(total_slides):
            # Update progress bar dan status
            progress_percentage = (i + 1) / total_slides
            progress_bar.progress(progress_percentage)
            status_text.text(f"Memproses slide {i + 1} dari {total_slides}...")

            time.sleep(1) # Beri waktu agar slide sempat di-render sepenuhnya

            # Ambil tangkapan layar dari elemen slide sebagai data biner PNG
            png_data = slide_container.screenshot_as_png
            image_data_list.append(png_data)

            # Klik tombol "next" jika bukan slide terakhir
            if i < total_slides - 1:
                next_button = driver.find_element(By.CSS_SELECTOR, next_button_selector)
                driver.execute_script("arguments[0].click();", next_button)

        status_text.text("Semua slide telah diproses. Membuat file ZIP...")
        
        # Buat file ZIP di dalam memori
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for idx, data in enumerate(image_data_list):
                # Simpan setiap data gambar ke dalam file ZIP
                zipf.writestr(f"slide_{idx + 1:03d}.png", data)
        
        progress_bar.progress(1.0)
        status_text.success("File ZIP berhasil dibuat dan siap diunduh!")
        return zip_buffer.getvalue()

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.info("Pastikan URL yang dimasukkan adalah URL 'embed' Google Slides dan coba muat ulang halaman.")
        return None
    # Driver tidak ditutup secara eksplisit karena di-cache oleh Streamlit

# --- Antarmuka Pengguna Streamlit ---
st.set_page_config(page_title="Pengunduh Google Slides", layout="centered")

st.title("🖼️ Pengunduh Google Slides")
st.markdown("Tempelkan URL `embed` dari Google Slides di bawah ini untuk mengunduh semua slide sebagai gambar dalam satu file ZIP.")

# URL default dari percakapan kita
default_url = "https://docs.google.com/presentation/d/e/2PACX-1vRoiaaNbJnyMe-Z19h9h2wy24yJsX_rFhHn6_svn5VMKMDcl4yMsjMF3qoNUOO_Yg/embed?start=false&loop=false&delayms=3000&slide=id.p1"
url = st.text_input("URL Google Slides:", value=default_url)

if st.button("Unduh Slide"):
    if url:
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        with st.spinner("Harap tunggu, proses ini mungkin memakan beberapa menit..."):
            zip_data = download_and_zip_slides(url, progress_bar, status_text)

        if zip_data:
            st.download_button(
                label="📥 Unduh slides.zip",
                data=zip_data,
                file_name="downloaded_slides.zip",
                mime="application/zip"
            )
    else:
        st.warning("Harap masukkan URL Google Slides terlebih dahulu.")

