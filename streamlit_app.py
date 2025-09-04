import streamlit as st
import subprocess
import sys
import asyncio
from playwright.async_api import async_playwright
import os
import shutil
import zipfile
import io

# --- Konfigurasi ---
DOWNLOAD_DIR = "slides_output"
DEFAULT_URL = "https://docs.google.com/presentation/d/e/2PACX-1vRoiaaNbJnyMe-Z19h9h2wy24yJsX_rFhHn6_svn5VMKMDcl4yMsjMF3qoNUOO_Yg/pub?start=false&loop=false&delayms=3000"

# --- FUNGSI INSTALASI (BARU) ---
@st.cache_resource
def setup_playwright():
    """
    Menjalankan perintah 'playwright install' menggunakan subprocess.
    Menggunakan @st.cache_resource agar hanya dijalankan sekali per sesi.
    """
    st.write("Memeriksa dan menginstal browser untuk Playwright...")
    
    # Perintah yang akan dijalankan di terminal
    command = [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"]
    
    try:
        # Menjalankan perintah
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True # Ini akan memunculkan error jika perintah gagal
        )
        st.write("Output Instalasi:")
        st.code(process.stdout)
        if process.stderr:
            st.warning("Pesan Tambahan dari Instalasi:")
            st.code(process.stderr)
        st.success("Setup Playwright selesai!")
        return True
    except subprocess.CalledProcessError as e:
        st.error("Gagal menjalankan instalasi Playwright.")
        st.error("Error Output:")
        st.code(e.stderr)
        return False
    except Exception as e:
        st.error(f"Terjadi kesalahan yang tidak terduga saat instalasi: {e}")
        return False

# --- Fungsi Inti Playwright ---
async def download_slides_with_playwright(url, progress_bar, status_text):
    if os.path.exists(DOWNLOAD_DIR): shutil.rmtree(DOWNLOAD_DIR)
    os.makedirs(DOWNLOAD_DIR)
    slide_paths = []
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"])
        page = await browser.new_page()
        try:
            status_text.text("Membuka URL...")
            await page.goto(url, wait_until="networkidle", timeout=60000)
            await page.wait_for_timeout(3000)

            paginator_text = await page.inner_text('.punch-viewer-paginator-page-count')
            total_slides = int(paginator_text)
            
            for i in range(total_slides):
                slide_number = i + 1
                progress_bar.progress(slide_number / total_slides)
                status_text.text(f"Mengambil gambar slide {slide_number}/{total_slides}...")
                
                slide_element = await page.query_selector(".punch-viewer-content")
                path = os.path.join(DOWNLOAD_DIR, f"slide_{slide_number:03d}.png")
                await slide_element.screenshot(path=path)
                slide_paths.append(path)
                
                if i < total_slides - 1:
                    await page.click(".punch-viewer-navbar-next")
                    await page.wait_for_timeout(1000)
        finally:
            await browser.close()
    return slide_paths

def create_zip_file(file_paths):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in file_paths:
            zf.write(file, os.path.basename(file))
    return zip_buffer.getvalue()

# --- Antarmuka Streamlit ---
st.set_page_config(page_title="Pengunduh Google Slides", layout="centered")
st.title("🚀 Pengunduh Google Slides")

# Menjalankan setup di awal
setup_success = setup_playwright()

if setup_success:
    st.info("Browser sudah siap. Silakan masukkan URL Google Slides Anda.")
    url_input = st.text_input("URL Google Slides", value=DEFAULT_URL)

    if st.button("Unduh Semua Slide", type="primary"):
        if url_input:
            final_url = url_input.split('/embed')[0] + '/pub?start=false&loop=false&delayms=3000'
            progress_bar = st.progress(0.0)
            status_text = st.empty()
            
            with st.spinner("Proses unduh sedang berjalan..."):
                slide_paths = asyncio.run(download_slides_with_playwright(final_url, progress_bar, status_text))
            
            if slide_paths:
                zip_data = create_zip_file(slide_paths)
                st.download_button("📥 Unduh slides.zip", zip_data, "downloaded_slides.zip", "application/zip")
                shutil.rmtree(DOWNLOAD_DIR)
else:
    st.error("Aplikasi tidak bisa berjalan karena setup gagal. Silakan periksa log di atas.")

