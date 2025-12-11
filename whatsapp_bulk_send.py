"""
whatsapp_bulk_send.py
Sends WhatsApp messages (via WhatsApp Web) to numbers listed in an Excel file.
No API keys required.

Notes:
- Put your message template below. Use {name} to personalize when 'name' column exists.
- Use responsibly and avoid spamming.
"""

import time
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Set Chrome executable path if needed
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    # Persistent login (so WhatsApp Web stays logged in)
    options.add_argument(r"--user-data-dir=C:\chrome-whatsapp-profile")

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        driver.get("https://web.whatsapp.com")
        return driver
    except Exception as e:
        print("[ERROR] Failed to start Chrome:", e)
        return None

# === CONFIG ===
EXCEL_FILE = "contacts.xlsx"     # path to your excel
SHEET_NAME = "Sheet1"            # sheet name
PHONE_COLUMN = "number"          # column with phone numbers
NAME_COLUMN = "name"             # optional name column (set to None if none)
MESSAGE_TEMPLATE = "Assalamualaikum {name},\nThis is a test message sent from my app. Please ignore."  # use {name} if you have names
DELAY_BETWEEN_MESSAGES = 6      # seconds to wait after sending each message (adjust if needed)
HEADLESS = False                # if True, runs Chrome headless (not recommended; better to see QR scan)
# =================

def clean_phone(number_str: str) -> str:
    """Normalize phone number: remove spaces, plus signs and non-digit chars."""
    if pd.isna(number_str):
        return ""
    s = str(number_str)
    s = ''.join(ch for ch in s if ch.isdigit())
    return s

def build_message(row):
    """Return URL-encoded message, handling optional name personalization."""
    if NAME_COLUMN and NAME_COLUMN in row and not pd.isna(row[NAME_COLUMN]):
        msg = MESSAGE_TEMPLATE.format(name=row[NAME_COLUMN])
    else:
        # If template contains {name} but name missing, remove placeholder gracefully
        msg = MESSAGE_TEMPLATE
        try:
            msg = msg.format(name="")
        except Exception:
            pass
    return urllib.parse.quote(msg)

def wait_for_login(driver, timeout=60):
    """Wait until WhatsApp Web is ready after QR scan by checking for the presence of the search box or chat pane."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
        )
        return True
    except TimeoutException:
        return False

def send_message_to_number(driver, phone, encoded_message):
    url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"
    driver.get(url)

    try:
        # Wait for chat box to load
        chat_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            )
        )

        time.sleep(1)

        # Click the chat box to activate it
        chat_box.click()
        time.sleep(0.5)

        # Press ENTER to send the message
        chat_box.send_keys(Keys.ENTER)
        time.sleep(1)

        # If ENTER fails, click the green send button
        try:
            send_button = driver.find_element(
                By.XPATH, "//span[@data-icon='send']"
            )
            send_button.click()
        except:
            pass  # ENTER already worked

        return True

    except Exception as e:
        print(f"[!] Failed to send to {phone}: {e}")
        return False


def main():
    # Load contacts
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    if PHONE_COLUMN not in df.columns:
        print(f"[ERROR] Excel file must contain a '{PHONE_COLUMN}' column")
        return

# Clean numbers and drop blanks
    # df[PHONE_COLUMN] = df[PHONE_COLUMN].apply(clean_phone)
    # df = df[df[PHONE_COLUMN] != ""].drop_duplicates(subset=[PHONE_COLUMN]).reset_index(drop=True)
    # if df.empty:
    #     print("[ERROR] No valid phone numbers found in the Excel file.")
    #     return


def normalize_number(number: str) -> str:
    number = number.strip().replace(" ", "").replace("-", "")

    # With country code already (92xxxxxxxxxx)
    if number.startswith("92") and len(number) == 12:
        return number

    # Starts with 03 (Pakistani format)
    if number.startswith("03") and len(number) == 11:
        return "92" + number[1:]   # replace 0 with 92

    # Starts with 0 but not 03
    if number.startswith("0") and len(number) >= 10:
        return "92" + number[1:]

    # Local number like 3xxxxxxxxx
    if number.startswith("3") and len(number) == 10:
        return "92" + number

    return number
    # Setup Chrome
    options = webdriver.ChromeOptions()

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # set Chrome executable path (VERY IMPORTANT)
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    # persistent login (so WhatsApp Web stays logged in)
    options.add_argument(r"--user-data-dir=C:\chrome-whatsapp-profile")

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
    except Exception as e:
        print("[ERROR] Failed to start Chrome:", e)
        return

    # Start driver
    print("Opening WhatsApp Web. Please scan the QR code if not already logged in.")
    driver.get("https://web.whatsapp.com")
    if not wait_for_login(driver, timeout=60):
        print("[ERROR] WhatsApp Web not ready - please ensure you scanned the QR code and are connected.")
        driver.quit()
        return
    print("Logged into WhatsApp Web.")

    # Iterate rows
    for idx, row in df.iterrows():
        phone = row[PHONE_COLUMN]
        if phone == "":
            continue
        msg = build_message(row)
        print(f"Sending to {phone} ({idx+1}/{len(df)}) ...")
        ok = send_message_to_number(driver, phone, msg)
        if ok:
            print(" -> Sent (or at least message entered).")
        else:
            print(" -> Failed (see warning above).")
        time.sleep(DELAY_BETWEEN_MESSAGES)

    print("Done. Closing browser in 5 seconds.")
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    main()