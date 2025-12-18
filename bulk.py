"""
whatsapp_bulk_send.py
Sends WhatsApp messages via WhatsApp Web to contacts listed in an Excel file.
No API keys required.
"""

import time
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# === CONFIG ===
EXCEL_FILE = "contacts.xlsx"      # Excel file path
SHEET_NAME = "Sheet1"             # Sheet name
PHONE_COLUMN = "contacts"         # Column with phone numbers
NAME_COLUMN = "name"              # Optional name column
MESSAGE_TEMPLATE = "Assalamualaikum {name},\nThis is a test message sent from my app. Please ignore."
DELAY_BETWEEN_MESSAGES = 8        # Seconds to wait after sending each message

# === Chrome profile path for persistent login ===
CHROME_PROFILE_PATH = r"C:\chrome-whatsapp-profile"

# ---------------- Phone cleaning ----------------
def clean_phone(phone) -> str:
    if pd.isna(phone):
        return ""
    phone = str(phone).strip()
    if phone.endswith(".0"):
        phone = phone[:-2]
    return "".join(ch for ch in phone if ch.isdigit())

def normalize_contacts(contact: str) -> str:
    contact = contact.strip().replace(" ", "").replace("-", "")
    if contact.startswith("92") and len(contact) == 12:
        return contact
    if contact.startswith("03") and len(contact) == 11:
        return "92" + contact[1:]
    if contact.startswith("0") and len(contact) >= 10:
        return "92" + contact[1:]
    if contact.startswith("3") and len(contact) == 10:
        return "92" + contact
    return contact

def build_message(row):
    """Return message text with optional {name} replacement"""
    if NAME_COLUMN and NAME_COLUMN in row and not pd.isna(row[NAME_COLUMN]):
        return MESSAGE_TEMPLATE.format(name=row[NAME_COLUMN])
    else:
        try:
            return MESSAGE_TEMPLATE.format(name="")
        except:
            return MESSAGE_TEMPLATE

# ---------------- Selenium driver ----------------
def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={CHROME_PROFILE_PATH}")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        return driver
    except Exception as e:
        print("[ERROR] Could not start Chrome:", e)
        return None

def wait_for_login(driver, timeout=60):
    """Wait until WhatsApp Web is ready after QR scan"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//div[@title='Search input textbox' or @contenteditable='true']"))
        )
        return True
    except:
        return False

# ---------------- Send message ----------------
def send_message(driver, phone, message):
    """Send message to single contact"""
    driver.get(f"https://web.whatsapp.com/send?phone={phone}&app_absent=0")
    try:
        # Wait for chat box (correct element, not search bar)
        chat_box = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            )
        )
        time.sleep(1)
        chat_box.click()
        chat_box.send_keys(message)
        time.sleep(0.5)
        chat_box.send_keys(Keys.ENTER)
        time.sleep(2)
        return True
    except Exception as e:
        print(f"[FAILED] {phone}: {e}")
        return False

# ---------------- Main ----------------
def main():
    # Load Excel
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"[ERROR] Could not read Excel file: {e}")
        return

    if PHONE_COLUMN not in df.columns:
        print(f"[ERROR] Excel must contain '{PHONE_COLUMN}' column")
        return

    # Clean & normalize
    df[PHONE_COLUMN] = df[PHONE_COLUMN].apply(clean_phone).apply(normalize_contacts)
    df = df[df[PHONE_COLUMN].str.len() == 12].drop_duplicates(subset=[PHONE_COLUMN]).reset_index(drop=True)

    if df.empty:
        print("[ERROR] No valid phone numbers after cleaning")
        return

    print("📋 Numbers after cleaning:")
    print(df[PHONE_COLUMN].tolist())

    # Start Chrome
    driver = init_driver()
    if not driver:
        print("[ERROR] Could not start Chrome")
        return

    print("📱 Opening WhatsApp Web. Scan QR code if needed...")
    if not wait_for_login(driver, timeout=120):
        print("[ERROR] WhatsApp Web not ready. Login failed or timed out.")
        driver.quit()
        return
    print("✅ Logged into WhatsApp Web")

    # Send messages
    for idx, row in df.iterrows():
        phone = row[PHONE_COLUMN]
        message = build_message(row)
        print(f"➡ Sending to {phone} ({idx+1}/{len(df)})...")
        success = send_message(driver, phone, message)
        if success:
            print(" -> Sent")
        else:
            print(" -> Failed")
        time.sleep(DELAY_BETWEEN_MESSAGES)

    print("🎉 Done. Closing browser in 5 seconds.")
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    main()