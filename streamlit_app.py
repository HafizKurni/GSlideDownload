import os
import time
import zipfile
import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title="Google Slides Downloader",
    page_icon="📄",
    layout="centered"
)

# --- Main App UI ---
st.title("📄 Google Slides Downloader")
st.write("Paste the 'embed' URL of a public Google Slides presentation to download all slides as a ZIP file containing PNG images.")

# --- Configuration ---
# Default URL for the input box
DEFAULT_URL = "https://docs.google.com/presentation/d/e/2PACX-1vRoiaaNbJnyMe-Z19h9h2wy24yJsX_rFhHn6_svn5VMKMDcl4yMsjMF3qoNUOO_Yg/embed?start=false&loop=false&delayms=3000&slide=id.p1"
OUTPUT_FOLDER = "temp_slides"
ZIP_FILENAME = "downloaded_slides.zip"

# --- Selenium Functions ---

@st.cache_resource
def setup_driver():
    """Sets up the Chrome WebDriver for Streamlit Cloud."""
    # These options are crucial for running Chrome in a headless
    # (no GUI) environment like the one Streamlit Cloud uses.
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # This prevents unnecessary log messages from cluttering the console
    options.add_argument("--log-level=3") 
    
    # webdriver-manager automatically downloads and manages the driver
    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def download_slides(driver, url):
    """Navigates to the URL and downloads each slide as a PNG image."""
    
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    status_placeholder = st.empty()
    progress_bar = st.progress(0)
    
    status_placeholder.info("Navigating to Google Slides URL...")
    driver.get(url)
    
    wait = WebDriverWait(driver, 20)

    try:
        paginator_selector = (By.CSS_SELECTOR, ".punch-viewer-nav-details > div")
        next_button_selector = (By.CSS_SELECTOR, "[aria-label='Next']")
        slide_container_selector = (By.ID, "punch-viewer-content-container")
        
        status_placeholder.info("Waiting for presentation elements to load...")
        paginator = wait.until(EC.presence_of_element_located(paginator_selector))
        next_button = wait.until(EC.element_to_be_clickable(next_button_selector))
        slide_container = wait.until(EC.presence_of_element_located(slide_container_selector))
        
        total_slides_text = paginator.text
        total_slides = int(total_slides_text.split('/')[-1].strip())
        status_placeholder.info(f"Found {total_slides} slides to download.")
        
        for i in range(1, total_slides + 1):
            progress_bar.progress(i / total_slides)
            status_placeholder.info(f"Downloading slide {i} of {total_slides}...")
            
            time.sleep(0.5)
            
            file_path = os.path.join(OUTPUT_FOLDER, f"slide_{str(i).zfill(3)}.png")
            slide_container.screenshot(file_path)
            
            if i < total_slides:
                next_button.click()
        
        progress_bar.progress(1.0)
        status_placeholder.success("All slides downloaded successfully! Now creating ZIP file...")
        return True

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please check that the URL is a valid 'embed' link and the presentation is public.")
        return False

def zip_files():
    """Zips the downloaded PNG files."""
    if os.path.exists(ZIP_FILENAME):
        os.remove(ZIP_FILENAME)

    with zipfile.ZipFile(ZIP_FILENAME, 'w') as zipf:
        for root, _, files in os.walk(OUTPUT_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)

def cleanup():
    """Removes temporary files and folders."""
    if os.path.exists(ZIP_FILENAME):
        os.remove(ZIP_FILENAME)
    if os.path.exists(OUTPUT_FOLDER):
        for file in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, file))
        os.rmdir(OUTPUT_FOLDER)

# --- Streamlit UI and Logic ---
slides_url = st.text_input("Enter Google Slides Embed URL", DEFAULT_URL)

if st.button("Download Slides", type="primary"):
    if slides_url:
        driver = setup_driver()
        
        cleanup()
        
        download_success = download_slides(driver, slides_url)
        
        if download_success:
            zip_files()
            with open(ZIP_FILENAME, "rb") as fp:
                st.download_button(
                    label="Download ZIP",
                    data=fp,
                    file_name="slides.zip",
                    mime="application/zip",
                    on_click=cleanup
                )
    else:
        st.warning("Please enter a URL.")

