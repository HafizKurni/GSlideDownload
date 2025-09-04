import streamlit as st
import asyncio
from playwright.async_api import async_playwright
import os
import shutil
import zipfile
import io

# --- Konfigurasi ---
DOWNLOAD_DIR = "slides_output"
DEFAULT_URL = "https://docs.google.com/presentation/d/e/2PACX-1vRoiaaNbJnyMe-Z19h9h2wy24yJsX_rFhHn6_svn5VMKMDcl4yMsjMF3qoNUOO_Yg/pub?start=false&loop=false&delayms=3000"

# --- Fungsi Inti Playwright ---
async def download_slides_with_playwright(url, progress_bar, status_text):
    """
    Menggunakan Playwright untuk membuka slide, mengambil screenshot, dan menyimpannya.
    """
    # Membersihkan direktori lama jika ada dan membuat yang baru
    if os.path.exists(DOWNLOAD_DIR):
        shutil.rmtree(DOWNLOAD_DIR)
    os.makedirs(DOWNLOAD_DIR)

    slide_paths = []
    
    async with async_playwright() as p:
        # Menjalankan browser dengan argumen tambahan untuk kompatibilitas cloud
        browser = await p.chromium.launch(
            headless=True, 
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        ) 
        page = await browser.new_page()
        
        try:
            status_text.text("Membuka URL Google Slides...")
            await page.goto(url, wait_until="networkidle", timeout=60000)
            await page.wait_for_timeout(3000) # Tunggu sebentar agar slide pertama termuat

            # Mencari elemen slide dan tombol navigasi
            slide_viewer_selector = ".punch-viewer-content"
            next_button_selector = ".punch-viewer-navbar-next"
            
            # Cek apakah elemen penting ada
            if not await page.is_visible(slide_viewer_selector) or not await page.is_visible(next_button_selector):
                st.error("Gagal menemukan elemen slide atau tombol navigasi. Pastikan URL benar dan dalam mode 'pub'.")
                await browser.close()
                return None

            # Mencari jumlah total slide dari elemen paginator
            paginator_text = await page.inner_text('.punch-viewer-paginator-page-count')
            total_slides = int(paginator_text)
            
            status_text.text(f"Ditemukan {total_slides} slide. Memulai proses screenshot...")

            for i in range(total_slides):
                slide_number = i + 1
                
                # Update progress bar dan status
                progress_percentage = slide_number / total_slides
                progress_bar.progress(progress_percentage)
                status_text.text(f"Mengambil gambar slide {slide_number} dari {total_slides}...")
                
                # Ambil screenshot dari elemen slide
                slide_element = await page.query_selector(slide_viewer_selector)
                screenshot_path = os.path.join(DOWNLOAD_DIR, f"slide_{slide_number:03d}.png")
                await slide_element.screenshot(path=screenshot_path)
                slide_paths.append(screenshot_path)
                
                # Klik tombol next jika bukan slide terakhir
                if i < total_slides - 1:
                    await page.click(next_button_selector)
                    await page.wait_for_timeout(1000) # Tunggu slide berikutnya termuat

            status_text.text("Semua slide berhasil diambil!")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat proses scraping: {e}")
            return None
        finally:
            await browser.close()

    return slide_paths

def create_zip_file(file_paths):
    """
    Membuat file zip dari daftar path gambar.
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in file_paths:
            zip_file.write(file_path, os.path.basename(file_path))
    return zip_buffer.getvalue()

# --- Antarmuka Streamlit ---
st.set_page_config(page_title="Pengunduh Google Slides", layout="centered")

st.title("🚀 Pengunduh Google Slides")
st.write(
    "Tempelkan URL Google Slides Anda di bawah ini dan dapatkan semua slide "
    "sebagai file gambar dalam satu file ZIP."
)
st.info(
    "**Penting:** Gunakan URL yang berakhiran dengan `/pub` atau `/embed`. "
    "Aplikasi ini akan mencoba mengubahnya secara otomatis.",
    icon="ℹ️"
)

url_input = st.text_input("URL Google Slides", value=DEFAULT_URL)

if st.button("Unduh Semua Slide", type="primary"):
    if not url_input:
        st.warning("Silakan masukkan URL terlebih dahulu.")
    else:
        # Memastikan URL dalam format yang benar ('/pub' lebih stabil)
        if "/embed" in url_input:
            final_url = url_input.split('/embed')[0] + '/pub?start=false&loop=false&delayms=3000'
        elif "/pub" not in url_input:
             # Coba konversi jika formatnya umum
             base_url_part = url_input.split('/edit')[0]
             final_url = base_url_part + '/pub?start=false&loop=false&delayms=3000'
        else:
            final_url = url_input
        
        st.write(f"Mengakses URL: `{final_url}`")
        
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        with st.spinner("Harap tunggu, browser virtual sedang disiapkan..."):
            # Menjalankan fungsi async dari dalam Streamlit
            slide_paths = asyncio.run(download_slides_with_playwright(final_url, progress_bar, status_text))

        if slide_paths:
            status_text.text("Membuat file ZIP...")
            zip_data = create_zip_file(slide_paths)
            status_text.success("Selesai! File ZIP Anda siap diunduh.")
            
            st.download_button(
                label="📥 Unduh slides.zip",
                data=zip_data,
                file_name="downloaded_slides.zip",
                mime="application/zip",
            )
            # Membersihkan direktori setelah selesai
            shutil.rmtree(DOWNLOAD_DIR)
        else:
            status_text.error("Proses gagal. Silakan periksa URL dan coba lagi.")

