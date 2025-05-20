import sqlite3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
from selenium.webdriver.chrome.options import Options
import argparse
import logging
from datetime import datetime

# Logging Setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler('scraper.log'),
        logging.StreamHandler()
    ]
)

# Configuration
CONFIG = {
    'headless': True,
    'max_retries': 3,
    'page_load_timeout': 45,
    'implicit_wait': 10,
    'db_name': 'eudamed_devices.db',
    'backup_interval': 6  # hours
}

def init_db():
    """Initialize SQLite database with all 70 columns"""
    conn = sqlite3.connect(CONFIG['db_name'])
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS devices (
        udi_di TEXT PRIMARY KEY,
        manufacturer_version TEXT,
        manufacturer_last_update TEXT,
        manufacturer_name TEXT,
        manufacturer_srn TEXT,
        manufacturer_address TEXT,
        manufacturer_country TEXT,
        manufacturer_phone TEXT,
        manufacturer_email TEXT,
        basic_udi_version TEXT,
        basic_udi_last_update TEXT,
        applicable_legislation TEXT,
        basic_udi_code TEXT,
        is_kit TEXT,
        is_system_procedure TEXT,
        authorised_rep TEXT,
        special_device_type TEXT,
        risk_class TEXT,
        is_implantable TEXT,
        is_suture_staple TEXT,
        has_measuring_function TEXT,
        is_reusable_instrument TEXT,
        is_active_device TEXT,
        administers_medicinal_product TEXT,
        is_companion_diagnostic TEXT,
        near_patient_testing TEXT,
        patient_self_testing TEXT,
        professional_testing TEXT,
        is_reagent TEXT,
        is_instrument TEXT,
        device_model TEXT,
        device_name TEXT,
        contains_human_tissue TEXT,
        contains_animal_tissue TEXT,
        contains_microbial_substance TEXT,
        contains_medicinal_product TEXT,
        contains_blood_plasma_product TEXT,
        udi_di_version TEXT,
        udi_di_last_update TEXT,
        udi_di_code TEXT,
        status TEXT,
        secondary_udi_di TEXT,
        nomenclature_codes TEXT,
        trade_names TEXT,
        catalogue_number TEXT,
        direct_marking_di TEXT,
        unit_of_use_di TEXT,
        quantity TEXT,
        udi_pi_type TEXT,
        additional_description TEXT,
        additional_info_url TEXT,
        clinical_sizes TEXT,
        labelled_single_use TEXT,
        max_reuses TEXT,
        needs_sterilisation TEXT,
        labelled_sterile TEXT,
        contains_latex TEXT,
        storage_conditions TEXT,
        critical_warnings TEXT,
        reprocessed_single_use TEXT,
        non_medical_purpose TEXT,
        eu_market_state TEXT,
        contains_medicinal_product_2 TEXT,
        market_version TEXT,
        market_last_update TEXT,
        available_member_states TEXT,
        sscp_reference TEXT,
        sscp_revision TEXT,
        sscp_issue_date TEXT,
        certificate_numbers TEXT,
        last_scraped TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Create indexes for performance
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_udi ON devices(udi_di)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_manufacturer ON devices(manufacturer_name)')
    conn.commit()
    return conn

def init_webdriver():
    """Configure Chrome WebDriver"""
    chrome_options = Options()
    if CONFIG['headless']:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(CONFIG['implicit_wait'])
    return driver

def safe_extract(driver, xpath, is_date=False, default="N/A"):
    """Safely extract text from XPath with error handling"""
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )    
        text = element.text.strip()
        if is_date and "Last update date:" in text:
            return text.replace("Last update date:", "").strip()
        return text if text else default
    except:
        return default

def extract_certificates(driver):
    """Handle complex certificate data extraction"""
    try:
        certificates = []
        # First try expansion panels
        panels = driver.find_elements(
            By.XPATH, "//h2[text()='Certificates']/following-sibling::div//mat-expansion-panel")
        for panel in panels:
            header = panel.find_element(By.TAG_NAME, "mat-expansion-panel-header").text
            content = panel.find_element(By.CLASS_NAME, "mat-expansion-panel-content").text
            certificates.append(f"{header}: {content}")
        
        # Fallback to simple divs if no panels found
        if not certificates:
            divs = driver.find_elements(
                By.XPATH, "//h2[text()='Certificates']/following-sibling::div/div")
            certificates = [div.text for div in divs if div.text]
        
        return " | ".join(certificates) if certificates else "N/A"
    except:
        return "N/A"

def process_device_page(driver, conn, current_page, device_num):
    """Process a single device detail page"""
    cursor = conn.cursor()
    try:
        data = {
            # Manufacturer Details (1-9)
            'udi_di': safe_extract(driver, "//dt[text()='UDI-DI code / Issuing entity']/following-sibling::dd/div"),
            'manufacturer_version': safe_extract(driver, "(//ul[@id='versionStatus']/li/strong)[1]"),
            'manufacturer_last_update': safe_extract(driver, "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[1]", True),
            # ... [All 70 fields here - truncated for brevity]
            'certificate_numbers': extract_certificates(driver)
        }
        
        # Insert data
        placeholders = ', '.join([':' + key for key in data.keys()])
        columns = ', '.join(data.keys())
        sql = f"INSERT OR REPLACE INTO devices ({columns}) VALUES ({placeholders})"
        cursor.execute(sql, data)
        conn.commit()
        
        logging.info(f"Page {current_page} | Device {device_num} | UDI: {data['udi_di'][:15]}... saved")
        return True
    except Exception as e:
        logging.error(f"Error processing device: {str(e)}")
        conn.rollback()
        return False
    finally:
        cursor.close()

def scrape_page(driver, conn, current_page):
    """Scrape all devices on a single page"""
    try:
        # Wait for table to load
        WebDriverWait(driver, CONFIG['page_load_timeout']).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
        
        devices = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        for i, device in enumerate(devices, 1):
            retries = 0
            success = False
            
            while retries < CONFIG['max_retries'] and not success:
                try:
                    # Open device details
                    view_btn = device.find_element(By.XPATH, ".//button[@title='View detail']")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", view_btn)
                    view_btn.click()
                    
                    # Process details
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='mb-5']")))
                    success = process_device_page(driver, conn, current_page, i)
                    
                    # Go back to list
                    driver.back()
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
                    
                except Exception as e:
                    retries += 1
                    logging.warning(f"Attempt {retries}/{CONFIG['max_retries']} failed: {str(e)}")
                    if retries >= CONFIG['max_retries']:
                        logging.error(f"Skipping device {i} on page {current_page} after {retries} retries")
                    time.sleep(5)
                    driver.refresh()
        
        return True
    except Exception as e:
        logging.error(f"Fatal error scraping page {current_page}: {str(e)}")
        return False

def main(start_page=1, end_page=2208):
    """Main scraping function"""
    conn = init_db()
    driver = init_webdriver()
    
    try:
        # Navigate to EUDAMED
        driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device?submitted=true")
        logging.info("EUDAMED loaded successfully")
        
        # Set 50 items per page
        dropdown = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "p-dropdown")))
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
        dropdown.click()
        
        option = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-label='50']")))
        option.click()
        
        # Verify 50 items loaded
        WebDriverWait(driver, 30).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "tbody > tr")) == 50)
        logging.info("50 items per page configured")
        
        # Main scraping loop
        for page in range(start_page, end_page + 1):
            logging.info(f"Starting page {page}/{end_page}")
            page_success = False
            page_retries = 0
            
            while not page_success and page_retries < CONFIG['max_retries']:
                try:
                    # Navigate to target page
                    if page > 1:
                        page_btn = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, f"//button[@aria-label='Page number {page}']")))
                        driver.execute_script("arguments[0].scrollIntoView(true);", page_btn)
                        page_btn.click()
                        WebDriverWait(driver, 30).until(
                            EC.staleness_of(page_btn))
                    
                    # Scrape current page
                    page_success = scrape_page(driver, conn, page)
                    
                except Exception as e:
                    page_retries += 1
                    logging.error(f"Page {page} attempt {page_retries} failed: {str(e)}")
                    time.sleep(10)
                    if page_retries >= CONFIG['max_retries']:
                        logging.critical(f"Permanent failure on page {page}")
                        break
                    driver.refresh()
            
            if not page_success:
                logging.error(f"Skipping page {page} after {page_retries} retries")
                continue
            
            # Backup every 6 hours
            if page % 50 == 0:  # Adjust based on scraping speed
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                backup_file = f"backups/eudamed_backup_{timestamp}.db"
                conn.execute(f"VACUUM INTO '{backup_file}'")
                logging.info(f"Database backup created: {backup_file}")
    
    finally:
        driver.quit()
        conn.close()
        logging.info("Scraping completed" if page == end_page else "Scraping interrupted")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-page", type=int, default=1, help="Page number to start from")
    parser.add_argument("--end-page", type=int, default=2208, help="Last page to scrape")
    parser.add_argument("--test-run", action="store_true", help="Run only 5 pages for testing")
    args = parser.parse_args()
    
    if args.test_run:
        CONFIG['headless'] = False
        args.end_page = min(5, args.end_page)
        logging.info(f"TEST MODE: Scraping pages {args.start_page}-{args.end_page}")
    
    main(start_page=args.start_page, end_page=args.end_page)