from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time

# Browser-Setup
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Maximiert das Browserfenster
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

# WebDriver initialisieren
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 40)

print("Initializing Chrome WebDriver and maximizing the browser window...")
driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device?submitted=true")

# Dropdown auf "50 items" setzen
print("Setting '50 items per page'...")
dropdown = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "p-dropdown")))
driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
dropdown.click()

dropdown_option = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-label='50']")))
driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_option)
dropdown_option.click()

# Warten bis 50 Zeilen geladen sind
def check_50_rows_loaded(driver):
    try:
        table = driver.find_element(By.TAG_NAME, "p-table")
        rows = table.find_elements(By.CSS_SELECTOR, "tbody > tr")
        return len(rows) == 50
    except:
        return False

wait.until(check_50_rows_loaded)
print("✅ Page with 50 items loaded.")
time.sleep(5)

# Beispielseiten (du kannst das beliebig erweitern)
pages_to_visit = [5, 7, 9, 11, 13, 15, 17, 19, 21]  # Beispielseiten, die du besuchen möchtest

for page in pages_to_visit:
    print(f"\n🔄 Navigating to Page {page}...")

    try:
        page_xpath = f"//button[contains(@aria-label, 'Page number {page} ')]"
        page_button = wait.until(EC.element_to_be_clickable((By.XPATH, page_xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", page_button)
        page_button.click()

        print(f"✅ Page {page} loaded.")
        time.sleep(3)

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        print(f"📄 {len(rows)} rows found.")

        if not rows:
            print("⚠️ No rows found on this page.")
            continue

        # Hier kannst du die Details für die erste Zeile extrahieren

    except Exception as e:
        print(f"❌ Error while processing page {page}: {e}")

# Fehlerbehandlung und benutzerdefinierte Eingaben
retry_iteration = False  # Initialer Wert

# Beispielseiten (pagesToVisit)
pages_to_visit = [5, 7, 9, 11, 13, 15]

for page in pages_to_visit:
    while True:
        try:
            page_xpath = f"//button[contains(@aria-label, 'Page number {page} ')]"
            print(f"\n🔄 Navigating to Page {page}...")

            # Warten, bis der Button klickbar ist, und klicken
            page_button = wait.until(EC.element_to_be_clickable((By.XPATH, page_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", page_button)
            page_button.click()

            print(f"✅ Successfully navigated to Page {page}")

            # Warte einige Sekunden, um die Seite vollständig zu laden
            time.sleep(3)

            break  # Erfolgreich navigiert, verlasse die Schleife

        except NoSuchElementException as ex:
            print(f"❌ An error occurred: {ex}")
            
            # Benutzeraufforderung zur Fehlerbehandlung
            print("Choose an option:")
            print("[R] Retry this iteration")
            print("[S] Skip to next page")
            print("[E] Exit program")

            user_choice = input().strip().upper()
            if user_choice == "R":
                print("Retrying...")
                retry_iteration = True
            elif user_choice == "S":
                print("Skipping to the next page...")
                retry_iteration = False
                break  # Gehe zur nächsten Seite
            elif user_choice == "E":
                print("Exiting program...")
                driver.quit()
                exit()  # Beende das Programm
            else:
                print("Invalid choice. Skipping iteration.")
                retry_iteration = False
                break  # Gehe zur nächsten Seite

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
import time

# Browser-Setup
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Maximiert das Browserfenster
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

# WebDriver initialisieren
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 40)

print("Initializing Chrome WebDriver and maximizing the browser window...")
driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device?submitted=true")

# Erstellen einer Excel-Datei
print("Creating an Excel workbook to store the extracted data...")
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Device Data"

# Setzen der Header für die Excel-Datei
print("Setting headers for the Excel file...")
headers = [
    "UDI-DI", "Version", "Last Update Date", "Actor/Organisation name", "Actor ID/SRN",
    "Address", "Country", "Telephone number", "Email", "Version", "Last update date", 
    "Applicable legislation", "Basic UDI-DI/EUDAMED DI / Issuing entity", "Kit", 
    "System/Procedure which is a device in itself", "Authorised representative", 
    "Special device type", "Risk class", "Implantable", 
    "Is the device a suture, staple, dental filling, dental brace, tooth crown, screw, wedge, plate, wire, pin, clip or connector?",
    "Measuring function", "Reusable surgical instrument", "Active device", 
    "Device intended to administer and / or remove medicinal product", "Companion diagnostic", 
    "Near patient testing", "Patient self testing", "Professional testing", "Reagent", 
    "Instrument", "Device model", "Device name", 
    "Presence of human tissues and cells or their derivatives", 
    "Presence of animal tissues and cells or their derivatives", 
    "Presence of cells or substance of microbial origin", 
    "Presence of a substance which, if used separately, may be considered to be a medicinal product", 
    "Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma", 
    "Version", "Last update date", "UDI-DI code / Issuing entity", "Status", 
    "UDI-DI from another entity (secondary)", "Nomenclature code(s)", 
    "Name/Trade name(s)", "Reference / Catalogue number", "Direct marking DI", 
    "Unit of Use DI", "Quantity of device", "Type of UDI-PI", "Additional Product description", 
    "Additional information url", "Clinical sizes", "Labelled as single use", 
    "Maximum number of reuses", "Need for sterilisation before use", "Device labelled as sterile", 
    "Containing Latex", "Storage and handling conditions", "Critical warnings or contra-indications", 
    "Reprocessed single use device", "Intended purpose other than medical (Annex XVI)", 
    "Member state of the placing on the EU market of the device", 
    "Presence of a substance which, if used separately, may be considered to be a medicinal product", 
    "Version", "Last update date", "Member State where the device is or is to be made available", 
    "SS(C)P Reference number", "SS(C)P revision number", "Issue date", "Certificates numbers"
]

# Setzen der Header in der ersten Zeile
for col_num, header in enumerate(headers, start=1):
    worksheet.cell(row=1, column=col_num, value=header)

# Starten der Iteration über die Zeilen
excel_row_index = 2

# Beispielseiten (pagesToVisit)
pages_to_visit = [5, 7, 9, 11, 13, 15]  # Beispielseiten, die du besuchen möchtest

for page in pages_to_visit:
    print(f"\n🔄 Navigating to Page {page}...")

    try:
        page_xpath = f"//button[contains(@aria-label, 'Page number {page} ')]"
        page_button = wait.until(EC.element_to_be_clickable((By.XPATH, page_xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", page_button)
        page_button.click()

        print(f"✅ Page {page} loaded.")
        time.sleep(3)

        table_rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        print(f"📄 {len(table_rows)} rows found.")

        if not table_rows:
            print("⚠️ No rows found on this page.")
            continue

        for row in table_rows:
            try:
                # Klicke auf den "View Detail"-Button
                view_button = row.find_element(By.XPATH, ".//button[@title='View detail']")
                driver.execute_script("arguments[0].scrollIntoView(true);", view_button)
                view_button.click()

                # Warten auf die Detailseite
                wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='mb-5']")))
                print("✅ Detail page loaded.")

                # Hier kannst du die Details extrahieren und in das Excel-Dokument einfügen
                # Beispielhafte Extraktion von Daten:
                udi_di = extract_field("//h1[contains(text(), 'UDI-DI')]", "UDI-DI", wait)
                version = extract_field("//div[@class='version']", "Version", wait)

                worksheet.cell(row=excel_row_index, column=1, value=udi_di)
                worksheet.cell(row=excel_row_index, column=2, value=version)

                excel_row_index += 1
                driver.back()
                time.sleep(3)

            except NoSuchElementException:
                print("❌ 'View detail' button not found.")
            except Exception as e:
                print(f"❌ Error while processing page {page}: {e}")

    except Exception as e:
        print(f"❌ Failed to load page {page}: {e}")

# Excel-Datei speichern
print("Saving Excel file...")
workbook.save("eudamed_data.xlsx")


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time

# Browser-Setup
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Maximiert das Browserfenster
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

# WebDriver initialisieren
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 40)

print("Initializing Chrome WebDriver and maximizing the browser window...")
driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device?submitted=true")

# Funktion zum Extrahieren von Feldern
def extract_field(xpath, field_name, wait):
    try:
        field_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        field_text = field_element.text.strip()
        print(f"📌 {field_name}: {field_text}")
        return field_text
    except NoSuchElementException:
        print(f"⚠️ {field_name} not found.")
        return ""

# Warten auf die Detailseite
def wait_for_detail_page():
    try:
        accordion_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='mb-5']")))
        print("Details have loaded.")
        return accordion_elements
    except NoSuchElementException:
        print("❌ Detail page not loaded.")
        return None

# Version und Last Update Date extrahieren
version_xpath = "(//ul[@id='versionStatus']/li/strong)[1]"
version_text = extract_field(version_xpath, "Version", wait)

last_update_xpath = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[1]"
last_update_text = extract_field(last_update_xpath, "Last Update Date", wait)

# Actor/Organisation Name extrahieren
actor_name_xpath = "//dt[text()='Actor/Organisation name']/following-sibling::dd/div"
actor_name_text = extract_field(actor_name_xpath, "Actor/Organisation Name", wait)

# Actor ID/SRN extrahieren
actor_id_xpath = "//dt[text()='Actor ID/SRN']/following-sibling::dd/div"
actor_id_text = extract_field(actor_id_xpath, "Actor ID/SRN", wait)

# Address extrahieren
address_xpath = "//dt[text()='Address']/following-sibling::dd/div"
address_text = extract_field(address_xpath, "Address", wait)

# Country extrahieren
country_xpath = "//dt[text()='Country']/following-sibling::dd/div"
country_text = extract_field(country_xpath, "Country", wait)

# Telephone number extrahieren
telephone_xpath = "//dt[text()='Telephone number']/following-sibling::dd/div"
telephone_text = extract_field(telephone_xpath, "Telephone Number", wait)

# Email extrahieren
email_xpath = "//dt[text()='Email']/following-sibling::dd/div"
email_text = extract_field(email_xpath, "Email", wait)

# Version 2 extrahieren (optional)
version_xpath2 = "(//ul[@id='versionStatus']/li/strong)[2]"
version_text2 = extract_field(version_xpath2, "Version 2", wait)

# Last Update Date 2 extrahieren
last_update_xpath2 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[2]"
last_update_text2 = extract_field(last_update_xpath2, "Last Update Date 2", wait)

# Applicable Legislation extrahieren
legislation_xpath = "//dt[contains(text(), 'Applicable legislation')]/following-sibling::dd/div"
legislation_text = extract_field(legislation_xpath, "Applicable Legislation", wait)

# Basic UDI-DI/EUDAMED DI / Issuing Entity extrahieren
udi_xpath = "//dt[contains(text(), 'UDI-DI/EUDAMED')]/following-sibling::dd/div"
udi_text = extract_field(udi_xpath, "Basic UDI-DI/EUDAMED DI / Issuing Entity", wait)

# Kit extrahieren
kit_xpath = "//dt[contains(text(), 'Kit')]/following-sibling::dd/div"
kit_text = extract_field(kit_xpath, "Kit", wait)

# System/Procedure extrahieren
system_procedure_xpath = "//dt[contains(text(), 'System')]/following-sibling::dd/div"
system_procedure_text = extract_field(system_procedure_xpath, "System/Procedure which is a device in itself", wait)

# Authorised Representative extrahieren
authorised_rep_xpath = "//dt[contains(text(), 'Authorised representative')]/following-sibling::dd/div"
authorised_rep_text = extract_field(authorised_rep_xpath, "Authorised Representative", wait)

# Ausgabe der extrahierten Daten
print("Extracted Data:")
print(f"Version: {version_text}")
print(f"Last Update Date: {last_update_text}")
print(f"Actor/Organisation Name: {actor_name_text}")
print(f"Actor ID/SRN: {actor_id_text}")
print(f"Address: {address_text}")
print(f"Country: {country_text}")
print(f"Telephone Number: {telephone_text}")
print(f"Email: {email_text}")
print(f"Version 2: {version_text2}")
print(f"Last Update Date 2: {last_update_text2}")
print(f"Applicable Legislation: {legislation_text}")
print(f"Basic UDI-DI/EUDAMED DI / Issuing Entity: {udi_text}")
print(f"Kit: {kit_text}")
print(f"System/Procedure which is a device in itself: {system_procedure_text}")
print(f"Authorised Representative: {authorised_rep_text}")

# Excel-Datei speichern
# workbook.save("eudamed_data.xlsx")


from selenium.common.exceptions import NoSuchElementException

# Extract Special Device Type
spec_dev_type_element = "//dt[contains(text(), 'Special device Type')]/following-sibling::dd/div"
spec_dev_type_text = ""
try:
    spec_dev_type_text = driver.find_element(By.XPATH, spec_dev_type_element).text
except NoSuchElementException:
    print("Special device type not found. Leaving it empty.")
print("Special device type: " + spec_dev_type_text)

# Extract Risk Class
risk_class_element = "//dt[contains(text(), 'Risk class')]/following-sibling::dd/div"
risk_class = ""
try:
    risk_class = driver.find_element(By.XPATH, risk_class_element).text
except NoSuchElementException:
    print("Risk class not found. Leaving it empty.")
print("Risk Class: " + risk_class)

# Extract Implantable
implantable_element = "//dt[contains(text(), 'Implantable')]/following-sibling::dd/div"
implantable = ""
try:
    implantable = driver.find_element(By.XPATH, implantable_element).text
except NoSuchElementException:
    print("Implantable not found. Leaving it empty.")
print("Implantable: " + implantable)

# Extract Suture/Staple Device
suture_element = "//dt[contains(text(), 'Is the device a suture, ')]/following-sibling::dd/div"
suture_device = ""
try:
    suture_device = driver.find_element(By.XPATH, suture_element).text
except NoSuchElementException:
    print("Suture device status not found. Leaving it empty.")
print("Is the device a suture/staple/etc: " + suture_device)

# Extract Measuring Function
measuring_function_element = "//dt[contains(text(), 'Measuring function')]/following-sibling::dd/div"
measuring_function = ""
try:
    measuring_function = driver.find_element(By.XPATH, measuring_function_element).text.strip()
except NoSuchElementException:
    print("Measuring Function not found. Leaving it empty.")
print("Measuring Function: " + measuring_function)

# Extract Reusable Surgical Instrument
reusable_instrument_element = "//dt[contains(text(), 'Reusable surgical instrument')]/following-sibling::dd/div"
reusable_instrument = ""
try:
    reusable_instrument = driver.find_element(By.XPATH, reusable_instrument_element).text.strip()
except NoSuchElementException:
    print("Reusable Surgical Instrument not found. Leaving it empty.")
print("Reusable Surgical Instrument: " + reusable_instrument)

# Extract Active Device
active_device_element = "//dt[contains(text(), 'Active device')]/following-sibling::dd/div"
active_device = ""
try:
    active_device = driver.find_element(By.XPATH, active_device_element).text.strip()
except NoSuchElementException:
    print("Active Device not found. Leaving it empty.")
print("Active Device: " + active_device)

# Extract Device Intended to Administer Medicinal Product
admin_device_element = "//dt[contains(text(), 'Device intended to administer and / or remove medicinal product')]/following-sibling::dd/div"
admin_device = ""
try:
    admin_device = driver.find_element(By.XPATH, admin_device_element).text
except NoSuchElementException:
    print("Device Intended to Administer Medicinal Product not found. Leaving it empty.")
print("Device Intended to Administer Medicinal Product: " + admin_device)

from selenium.common.exceptions import NoSuchElementException

# Extract Companion Diagnostic
comp_diag_element = "//dt[contains(text(), 'Companion diagnostic')]/following-sibling::dd/div"
comp_diag_text = ""
try:
    comp_diag_text = driver.find_element(By.XPATH, comp_diag_element).text
except NoSuchElementException:
    print("Companion diagnostic not found. Leaving it empty.")
print("Companion diagnostic: " + comp_diag_text)

# Extract Near Patient Testing
near_pat_test_element = "//dt[contains(text(), 'Near patient testing')]/following-sibling::dd/div"
near_pat_test_text = ""
try:
    near_pat_test_text = driver.find_element(By.XPATH, near_pat_test_element).text
except NoSuchElementException:
    print("Near patient testing not found. Leaving it empty.")
print("Near patient testing: " + near_pat_test_text)

# Extract Patient Self Testing
pat_self_test_element = "//dt[contains(text(), 'Patient self testing')]/following-sibling::dd/div"
pat_self_test_text = ""
try:
    pat_self_test_text = driver.find_element(By.XPATH, pat_self_test_element).text
except NoSuchElementException:
    print("Patient self testing not found. Leaving it empty.")
print("Patient self testing: " + pat_self_test_text)

# Extract Professional Testing
prof_test_element = "//dt[contains(text(), 'Professional testing')]/following-sibling::dd/div"
prof_test_text = ""
try:
    prof_test_text = driver.find_element(By.XPATH, prof_test_element).text
except NoSuchElementException:
    print("Professional testing not found. Leaving it empty.")
print("Professional testing: " + prof_test_text)

# Extract Reagent
reagent_element = "//dt[contains(text(), 'Reagent')]/following-sibling::dd/div"
reagent_text = ""
try:
    reagent_text = driver.find_element(By.XPATH, reagent_element).text
except NoSuchElementException:
    print("Reagent not found. Leaving it empty.")
print("Reagent: " + reagent_text)

# Extract Instrument
instrument_element = "//dt[contains(text(), 'Instrument')]/following-sibling::dd/div"
instrument_text = ""
try:
    instrument_text = driver.find_element(By.XPATH, instrument_element).text
except NoSuchElementException:
    print("Instrument not found. Leaving it empty.")
print("Instrument: " + instrument_text)

# Extract Device Model
device_model_element = "//dt[contains(text(), 'Device model')]/following-sibling::dd/div"
device_model_text = ""
try:
    device_model_text = driver.find_element(By.XPATH, device_model_element).text
except NoSuchElementException:
    print("Device model not found. Leaving it empty.")
print("Device model: " + device_model_text)

# Extract Device Name
device_name_element = wait.until(d => d.find_element(By.XPATH, "//dt[contains(text(), 'Device name')]/following-sibling::dd/div"))
device_name = device_name_element.text.strip()
print("Device Name: " + device_name)

# Tissues and Cells

# Extract Presence of Human Tissues and Cells or Their Derivatives
human_tissues_xpath = "//dt[text()='Presence of human tissues and cells or their derivatives']/following-sibling::dd/div"
presence_of_human_tissues = driver.find_element(By.XPATH, human_tissues_xpath).text
print("Presence of human tissues and cells or their derivatives: " + presence_of_human_tissues)

# Extract Presence of Animal Tissues and Cells or Their Derivatives
animal_tissues_xpath = "//dt[text()='Presence of animal tissues and cells or their derivatives']/following-sibling::dd/div"
presence_of_animal_tissues = driver.find_element(By.XPATH, animal_tissues_xpath).text
print("Presence of animal tissues and cells or their derivatives: " + presence_of_animal_tissues)

# Extract Presence of Cells or Substances of Microbial Origin
microbial_element = "//dt[contains(text(), 'Presence of cells or substances of microbial origin')]/following-sibling::dd/div"
microbial_text = ""
try:
    microbial_text = driver.find_element(By.XPATH, microbial_element).text
except NoSuchElementException:
    print("Presence of cells or substances of microbial origin not found. Leaving it empty.")
print("Presence of cells or substances of microbial origin: " + microbial_text)

# Information on Substances

# Extract Presence of a Substance Which, if Used Separately, May Be Considered to Be a Medicinal Product
medicinal_product_xpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product']/following-sibling::dd/div"
presence_of_medicinal_product = driver.find_element(By.XPATH, medicinal_product_xpath).text
print("Presence of a substance which, if used separately, may be considered to be a medicinal product: " + presence_of_medicinal_product)

# Extract Presence of a Substance Which, if Used Separately, May Be Considered to Be a Medicinal Product Derived from Human Blood or Human Plasma
blood_plasma_product_xpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma']/following-sibling::dd/div"
presence_of_blood_plasma_product = driver.find_element(By.XPATH, blood_plasma_product_xpath).text
print("Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma: " + presence_of_blood_plasma_product)

# UDI-DI Details

# Extract Version 1 (Current)
version_xpath_3 = "(//ul[@id='versionStatus']/li/strong)[3]"
version_text_3 = ""
try:
    version_text_3 = driver.find_element(By.XPATH, version_xpath_3).text
except NoSuchElementException:
    print("Version 3 not found. Leaving it empty.")
print("Version: " + version_text_3)

# Extract Last Update Date
last_update_xpath_3 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[3]"
last_update_text_3 = ""
try:
    last_update_text_3 = driver.find_element(By.XPATH, last_update_xpath_3).text.replace("Last update date: ", "").strip()
except NoSuchElementException:
    print("Last Update Date 3 not found. Leaving it empty.")
print("Last Update Date: " + last_update_text_3)

from selenium.common.exceptions import NoSuchElementException

# Extract the UDI-DI code / Issuing entity
udi_di_xpath = "//dt[text()='UDI-DI code / Issuing entity']/following-sibling::dd/div"
udi_di = driver.find_element(By.XPATH, udi_di_xpath).text
print("UDI-DI code / Issuing entity: " + udi_di)

# Extract the Status
status_xpath = "//dt[text()='Status']/following-sibling::dd/div"
status = driver.find_element(By.XPATH, status_xpath).text
print("Status: " + status)

# Extract the UDI-DI from another entity (secondary)
secondary_udi_xpath = "//dt[text()='UDI-DI from another entity (secondary)']/following-sibling::dd/div"
secondary_udi = ""
try:
    secondary_udi = driver.find_element(By.XPATH, secondary_udi_xpath).text
except NoSuchElementException:
    print("UDI-DI from another entity (secondary) not found. Leaving it empty.")
print("UDI-DI from another entity (secondary): " + secondary_udi)

# Extract the Nomenclature code(s)
nomenclature_code_xpath = "//dt[text()='Nomenclature code(s)']/following-sibling::dd/div"
nomenclature_code = driver.find_element(By.XPATH, nomenclature_code_xpath).text
print("Nomenclature code(s): " + nomenclature_code)

# Extract the Name/Trade name(s)
trade_name_xpath = "//dt[text()='Name/Trade name(s)']/following-sibling::dd/div"
trade_name = driver.find_element(By.XPATH, trade_name_xpath).text
print("Name/Trade name(s): " + trade_name)

# Extract the Reference / Catalogue number
catalogue_number_xpath = "//dt[text()='Reference / Catalogue number']/following-sibling::dd/div"
catalogue_number = driver.find_element(By.XPATH, catalogue_number_xpath).text
print("Reference / Catalogue number: " + catalogue_number)

# Extract the Direct marking DI
direct_marking_xpath = "//dt[text()='Direct marking DI']/following-sibling::dd/div"
direct_marking = ""
try:
    direct_marking = driver.find_element(By.XPATH, direct_marking_xpath).text
except NoSuchElementException:
    print("Direct marking DI not found. Leaving it empty.")
print("Direct marking DI: " + direct_marking)

# Extract Unit of Use
unit_of_use_element = "//dt[contains(text(), 'Unit of use')]/following-sibling::dd/div"
unit_of_use_text = ""
try:
    unit_of_use_text = driver.find_element(By.XPATH, unit_of_use_element).text
except NoSuchElementException:
    print("Unit of Use not found. Leaving it empty.")
print("Unit of Use: " + unit_of_use_text)

# Extract the Quantity of device
quantity_xpath = "//dt[text()='Quantity of device']/following-sibling::dd/div"
quantity = ""
try:
    quantity = driver.find_element(By.XPATH, quantity_xpath).text
except NoSuchElementException:
    print("Quantity of device not found. Leaving it empty.")
print("Quantity of device: " + quantity)

# Extract the Type of UDI-PI
udi_pi_xpath = "//dt[text()='Type of UDI-PI']/following-sibling::dd/div"
udi_pi = ""
try:
    udi_pi = driver.find_element(By.XPATH, udi_pi_xpath).text
except NoSuchElementException:
    print("Type of UDI-PI not found. Leaving it empty.")
print("Type of UDI-PI: " + udi_pi)

# Extract the Additional Product description
additional_description_xpath = "//dt[text()='Additional Product description']/following-sibling::dd/div"
additional_description = ""
try:
    additional_description = driver.find_element(By.XPATH, additional_description_xpath).text
except NoSuchElementException:
    print("Additional Product description not found. Leaving it empty.")
print("Additional Product description: " + additional_description)

# Extract the Additional information url
info_url_xpath = "//dt[text()='Additional information url']/following-sibling::dd/div"
info_url = ""
try:
    info_url = driver.find_element(By.XPATH, info_url_xpath).text
except NoSuchElementException:
    print("Additional information url not found. Leaving it empty.")
print("Additional information url: " + info_url)

# Extract the Clinical sizes
clinical_sizes_xpath = "//dt[text()='Clinical sizes']/following-sibling::dd/div"
clinical_sizes = ""
try:
    clinical_sizes = driver.find_element(By.XPATH, clinical_sizes_xpath).text
except NoSuchElementException:
    print("Clinical sizes not found. Leaving it empty.")
print("Clinical sizes: " + clinical_sizes)

# Extract the Labelled as single use
single_use_xpath = "//dt[text()='Labelled as single use']/following-sibling::dd/div"
single_use = ""
try:
    single_use = driver.find_element(By.XPATH, single_use_xpath).text
except NoSuchElementException:
    print("Labelled as single use not found. Leaving it empty.")
print("Labelled as single use: " + single_use)

# Extract the Maximum number of reuses
max_no_reuses_element = "//dt[text()='Maximum number of reuses']/following-sibling::dd/div"
max_no_reuses_text = ""
try:
    max_no_reuses_text = driver.find_element(By.XPATH, max_no_reuses_element).text
except NoSuchElementException:
    print("Maximum number of reuses not found. Leaving it empty.")
print("Maximum number of reuses: " + max_no_reuses_text)


# Extract the Need for sterilisation before use
sterilisation_xpath = "//dt[text()='Need for sterilisation before use']/following-sibling::dd/div"
sterilisation = ""
try:
    sterilisation = driver.find_element(By.XPATH, sterilisation_xpath).text
except NoSuchElementException:
    print("Need for sterilisation before use not found. Leaving it empty.")
print("Need for sterilisation before use: " + sterilisation)

# Extract the Device labelled as sterile
sterile_xpath = "//dt[text()='Device labelled as sterile']/following-sibling::dd/div"
sterile = ""
try:
    sterile = driver.find_element(By.XPATH, sterile_xpath).text
except NoSuchElementException:
    print("Device labelled as sterile not found. Leaving it empty.")
print("Device labelled as sterile: " + sterile)

# Extract the Containing Latex
latex_xpath = "//dt[text()='Containing Latex']/following-sibling::dd/div"
latex = ""
try:
    latex = driver.find_element(By.XPATH, latex_xpath).text
except NoSuchElementException:
    print("Containing Latex not found. Leaving it empty.")
print("Containing Latex: " + latex)

# Extract the Storage and handling conditions
handling_cond_element = "//dt[text()='Storage and handling conditions']/following-sibling::dd/div"
handling_cond_text = ""
try:
    handling_cond_text = driver.find_element(By.XPATH, handling_cond_element).text
except NoSuchElementException:
    print("Storage and handling conditions not found. Leaving it empty.")
print("Storage and handling conditions: " + handling_cond_text)

# Extract the Critical warnings or contra-indications
warnings_xpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd/div"
warnings = ""
try:
    warnings = driver.find_element(By.XPATH, warnings_xpath).text
except NoSuchElementException:
    print("Critical warnings or contra-indications not found. Leaving it empty.")
print("Critical warnings or contra-indications: " + warnings)

# Extract the Do not re-use
do_not_reuse_xpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd//li[text()='Do not re-use']"
do_not_reuse = ""
try:
    do_not_reuse = driver.find_element(By.XPATH, do_not_reuse_xpath).text
except NoSuchElementException:
    print("Do not re-use not found. Leaving it empty.")
print("Do not re-use: " + do_not_reuse)

# Extract the Reprocessed single use device
reprocessed_xpath = "//dt[contains(text(), 'Reprocessesed single use device')]/following-sibling::dd/div"
reprocessed = ""
try:
    reprocessed = driver.find_element(By.XPATH, reprocessed_xpath).text
except NoSuchElementException:
    print("Reprocessed single use device not found. Leaving it empty.")
print("Reprocessed single use device: " + reprocessed)

# Extract the Intended purpose other than medical (Annex XVI)
intended_purpose_xpath = "//dt[contains(text(), 'Intended purpose other than medical')]/following-sibling::dd/div"
intended_purpose = ""
try:
    intended_purpose = driver.find_element(By.XPATH, intended_purpose_xpath).text
except NoSuchElementException:
    print("Intended purpose other than medical (Annex XVI) not found. Leaving it empty.")
print("Intended purpose other than medical (Annex XVI): " + intended_purpose)

# Extract the Member state of the placing on the EU market of the device
member_state_xpath = "//dt[text()='Member state of the placing on the EU market of the device']/following-sibling::dd/div"
member_state = ""
try:
    member_state = driver.find_element(By.XPATH, member_state_xpath).text
except NoSuchElementException:
    print("Member state of the placing on the EU market of the device not found. Leaving it empty.")
print("Member state of the placing on the EU market of the device: " + member_state)

# Extract the Presence of a substance which, if used separately, may be considered to be a medicinal product
med_prod_element = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product']/following-sibling::dd/div"
med_prod_text = ""
try:
    med_prod_text = driver.find_element(By.XPATH, med_prod_element).text
except NoSuchElementException:
    print("Presence of a substance which, if used separately, may be considered to be a medicinal product not found. Leaving it empty.")
print("Presence of a substance which, if used separately, may be considered to be a medicinal product: " + med_prod_text)

# Market distribution
# Extract the Version 1 (Current)
version_xpath4 = "(//ul[@id='versionStatus']/li/strong)[4]"
version_text4 = ""
try:
    version_text4 = driver.find_element(By.XPATH, version_xpath4).text
except NoSuchElementException:
    print("Version 4 not found. Leaving it empty.")
print("Version: " + version_text4)

# Extract the Last update date
last_update_xpath4 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[4]"
last_update_text4 = ""
try:
    last_update_text4 = driver.find_element(By.XPATH, last_update_xpath4).text.replace("Last update date: ", "").strip()
except NoSuchElementException:
    print("Last Update Date 4 not found. Leaving it empty.")
print("Last Update Date: " + last_update_text4)

# Extract the Member State where the device is or is to be made available
member_state_xpath2 = "//dt[text()='Member State where the device is or is to be made available']/following-sibling::dd//ul"
member_state_availab = ""
try:
    member_state_availab = driver.find_element(By.XPATH, member_state_xpath2).text
except NoSuchElementException:
    print("Member State not found. Leaving it empty.")
print("Member State: " + member_state_availab)

# Extract the SS(C)P Reference number
ref_no_element = "//dt[text()='SS(C)P Reference number']/following-sibling::dd/div"
ref_no_text = ""
try:
    ref_no_text = driver.find_element(By.XPATH, ref_no_element).text
except NoSuchElementException:
    print("SS(C)P Reference number not found. Leaving it empty.")
print("SS(C)P Reference number: " + ref_no_text)

# Extract the SS(C)P revision number
rev_no_element = "//dt[text()='SS(C)P revision number']/following-sibling::dd/div"
rev_no_text = ""
try:
    rev_no_text = driver.find_element(By.XPATH, rev_no_element).text
except NoSuchElementException:
    print("SS(C)P revision number not found. Leaving it empty.")
print("SS(C)P revision number: " + rev_no_text)

# Extract the Issue date
issue_date_element = "//dt[text()='Issue date']/following-sibling::dd/div"
issue_date_text = ""
try:
    issue_date_text = driver.find_element(By.XPATH, issue_date_element).text
except NoSuchElementException:
    print("SS(C)P issue date not found. Leaving it empty.")
print("SS(C)P issue date: " + issue_date_text)


# Extract the Certificates numbers
certificate_no_element = "//h2[text()='Certificates']/following-sibling::div[1]//mat-expansion-panel-header"
certificate_no_element2 = "//h2[text()='Certificates']/following-sibling::div[1]//mat-expansion-panel/div/div/div"
certificate_no_text = ""

try:
    # Find all matching elements
    certificate_elements = driver.find_elements(By.XPATH, certificate_no_element)
    certificate_elements2 = driver.find_elements(By.XPATH, certificate_no_element2)

    if len(certificate_elements) > 0 or len(certificate_elements2) > 0:
        # Extract text from both sets of elements and concatenate them with " % "
        certificate_texts = [el.text for el in certificate_elements] + [el.text for el in certificate_elements2]
        certificate_no_text = "  %  ".join(certificate_texts) + " % "
    else:
        print("Certificates numbers not found. Leaving it empty.")
except NoSuchElementException:
    print("Certificates numbers not found. Leaving it empty.")

# Print the final output
print("Certificates numbers: " + certificate_no_text)

# Save extracted data to Excel
print(f"Saving data for UDI-DI: {udi_di}...")

# Assuming `worksheet` is already defined and `excel_row_index` is set to the correct row number
worksheet.cell(excel_row_index, 2).value = version_text
worksheet.cell(excel_row_index, 3).value = last_update_text
worksheet.cell(excel_row_index, 4).value = actor_name_text
worksheet.cell(excel_row_index, 5).value = actor_id_text
worksheet.cell(excel_row_index, 6).value = address_text
worksheet.cell(excel_row_index, 7).value = country_text
worksheet.cell(excel_row_index, 8).value = telephone_text
worksheet.cell(excel_row_index, 9).value = email_text

# Basic UDI-DI
worksheet.cell(excel_row_index, 10).value = version_text2
worksheet.cell(excel_row_index, 11).value = last_update_text2
worksheet.cell(excel_row_index, 12).value = applicable_legislation
worksheet.cell(excel_row_index, 13).value = udi_text_basic
worksheet.cell(excel_row_index, 14).value = kit_text
worksheet.cell(excel_row_index, 15).value = system_procedure
worksheet.cell(excel_row_index, 16).value = authorised_rep
worksheet.cell(excel_row_index, 17).value = spec_dev_type_text
worksheet.cell(excel_row_index, 18).value = risk_class
worksheet.cell(excel_row_index, 19).value = implantable
worksheet.cell(excel_row_index, 20).value = suture_device
worksheet.cell(excel_row_index, 21).value = measuring_function
worksheet.cell(excel_row_index, 22).value = reusable_instrument
worksheet.cell(excel_row_index, 23).value = active_device
worksheet.cell(excel_row_index, 24).value = admin_device
worksheet.cell(excel_row_index, 25).value = comp_diag_text
worksheet.cell(excel_row_index, 26).value = near_pat_test_text
worksheet.cell(excel_row_index, 27).value = pat_self_test_text
worksheet.cell(excel_row_index, 28).value = prof_test_text
worksheet.cell(excel_row_index, 29).value = reagent_text
worksheet.cell(excel_row_index, 30).value = instrument_text
worksheet.cell(excel_row_index, 31).value = device_model_text
worksheet.cell(excel_row_index, 32).value = device_name

# Tissues and cells
worksheet.cell(excel_row_index, 33).value = presence_of_human_tissues
worksheet.cell(excel_row_index, 34).value = presence_of_animal_tissues
worksheet.cell(excel_row_index, 35).value = microbial_text

# Information on Substances
worksheet.cell(excel_row_index, 36).value = presence_of_medicinal_product
worksheet.cell(excel_row_index, 37).value = presence_of_blood_plasma_product

# UDI-DI details
worksheet.cell(excel_row_index, 38).value = version_text3
worksheet.cell(excel_row_index, 39).value = last_update_text3
worksheet.cell(excel_row_index, 40).value = udi_di
worksheet.cell(excel_row_index, 41).value = status
worksheet.cell(excel_row_index, 42).value = secondary_udi
worksheet.cell(excel_row_index, 43).value = nomenclature_code
worksheet.cell(excel_row_index, 44).value = trade_name
worksheet.cell(excel_row_index, 45).value = catalogue_number
worksheet.cell(excel_row_index, 46).value = direct_marking
worksheet.cell(excel_row_index, 47).value = unit_of_use_text
worksheet.cell(excel_row_index, 48).value = quantity
worksheet.cell(excel_row_index, 49).value = udi_pi
worksheet.cell(excel_row_index, 50).value = additional_description
worksheet.cell(excel_row_index, 51).value = info_url
worksheet.cell(excel_row_index, 52).value = clinical_sizes
worksheet.cell(excel_row_index, 53).value = single_use
worksheet.cell(excel_row_index, 54).value = max_no_reuses_text
worksheet.cell(excel_row_index, 55).value = sterilisation
worksheet.cell(excel_row_index, 56).value = sterile
worksheet.cell(excel_row_index, 57).value = latex
worksheet.cell(excel_row_index, 58).value = handling_cond_text
worksheet.cell(excel_row_index, 59).value = warnings
worksheet.cell(excel_row_index, 60).value = reprocessed
worksheet.cell(excel_row_index, 61).value = intended_purpose
worksheet.cell(excel_row_index, 62).value = member_state
worksheet.cell(excel_row_index, 63).value = med_prod_text

# Market distribution
worksheet.cell(excel_row_index, 64).value = version_text4
worksheet.cell(excel_row_index, 65).value = last_update_text4
worksheet.cell(excel_row_index, 66).value = member_state_availab
worksheet.cell(excel_row_index, 67).value = ref_no_text
worksheet.cell(excel_row_index, 68).value = rev_no_text
worksheet.cell(excel_row_index, 69).value = issue_date_text
worksheet.cell(excel_row_index, 70).value = certificate_no_text

worksheet.cell(excel_row_index, 1).value = udi_di

print(f"*****************************************************************Data saved in row {excel_row_index}")
excel_row_index += 1

# Go back to the previous page
print("Navigating back to the previous page...")
driver.back()

# Save the Excel file
print("Saving the extracted data to an Excel file...")
workbook.save("Eudamed_Device_Data_2209.xlsx")

print(f"Data extraction for product No {i + 1}! Excel file saved as 'Eudamed_Device_Data_2209.xlsx'.")

# Wait for the table to reload
print("Waiting for the table to reload...")
time.sleep(5)  # Adjust as needed


from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time

def navigate_to_next_page(driver, current_page):
    wait = WebDriverWait(driver, 30)

    try:
        # XPath für die Schaltfläche "Nächste Seite"
        next_page_button_xpath = f"//button[@aria-label='Page number {current_page + 1} ']"
        next_page_button = wait.until(EC.element_to_be_clickable((By.XPATH, next_page_button_xpath)))
        
        # Scrollen und auf die Schaltfläche klicken
        driver.execute_script("arguments[0].scrollIntoView(true);", next_page_button)
        next_page_button.click()

        # Warten, bis die Tabelle aktualisiert wird
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "table tbody tr")) > 0)
        print(f"Page {current_page + 1} loaded.")
    
    except Exception as e:
        print(f"Error navigating to page {current_page + 1}: {e}")

def click_page(driver, page_number):
    try:
        # XPath für die Schaltfläche "Seite {page_number}"
        page_button = driver.find_element(By.XPATH, f"//button[contains(@aria-label, 'Page number {page_number}')]")
        page_button.click()
        print(f"Navigated to Page {page_number}")
    except NoSuchElementException:
        print(f"Page {page_number} button not found!")

# Beispiel: Verwendung der Funktionen
# Angenommen, `driver` ist der WebDriver und `current_page` ist die aktuelle Seite
current_page = 2208  # Beispielstartseite
driver = None  # Hier musst du den WebDriver initialisieren (z.B. `driver = webdriver.Chrome(options=options)`)

retry_iteration = False
total_pages = 22222  # Beispiel für eine große Anzahl von Seiten

try:
    for current_page in range(current_page, total_pages):
        print(f"Moving to page {current_page + 1}...")
        navigate_to_next_page(driver, current_page)
        
        # Warten, bis die Tabellenzeilen sichtbar sind
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "table tbody tr")))

except Exception as ex:
    print(f"An error occurred: {ex}")

    # Benutzeraufforderung zur Auswahl einer Aktion
    print("Choose an option:")
    print("[R] Retry this iteration")
    print("[S] Skip to next iteration")
    print("[E] Exit program")

    user_choice = input().strip().upper()
    if user_choice == "R":
        print("Retrying...")
    elif user_choice == "S":
        print("Skipping to next product...")
    elif user_choice == "E":
        print("Exiting program...")
        driver.quit()
    else:
        print("Invalid choice. Skipping iteration.")
