import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import time
import urllib.parse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import sqlite3
import datetime

# ---------------- Database Setup ----------------
DB_FILE = "whatsapp_contacts.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT UNIQUE,
            name TEXT,
            status TEXT,
            timestamp TEXT
        )
    ''')
    conn.commit()
    conn.close()

def add_contact(phone, name):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    try:
        c.execute("INSERT OR IGNORE INTO contacts (phone, name, status) VALUES (?, ?, ?)", 
                  (phone, name, "Pending"))
        conn.commit()
    except Exception as e:
        print("DB insert error:", e)
    conn.close()

def update_status(phone, status):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE contacts SET status=?, timestamp=? WHERE phone=?", (status, timestamp, phone))
    conn.commit()
    conn.close()

# ---------------------------------------------
#           WhatsApp Functions
# ---------------------------------------------

def clean_phone(number_str):
    if pd.isna(number_str):
        return ""
    s = str(number_str)
    return "".join(ch for ch in s if ch.isdigit())

def send_message_to_number(driver, phone, encoded_message, log_callback):
    url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"
    driver.get(url)

    try:
        chat_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            )
        )

        time.sleep(1)
        chat_box.click()
        time.sleep(0.5)

        chat_box.send_keys(Keys.ENTER)
        time.sleep(1)

        try:
            send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
            send_button.click()
        except:
            pass

        log_callback(f"✔ Sent to {phone}")
        return True

    except Exception as e:
        log_callback(f"❌ Failed: {phone} — {e}")
        return False

# ---------------------------------------------
#               Modern GUI
# ---------------------------------------------

class WhatsAppModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("WhatsApp Bulk Messenger — Modern Edition")
        self.geometry("780x700")

        ctk.set_appearance_mode("dark")     # "light" / "dark" / "system"
        ctk.set_default_color_theme("blue") # blue, dark-blue, green

        self.excel_path = ""

        # ---------- HEADER ----------
        header = ctk.CTkLabel(self, text="WhatsApp Bulk Messaging App", 
                              font=("Poppins", 28, "bold"))
        header.pack(pady=15)

        # ---------- FILE SELECT ----------
        file_frame = ctk.CTkFrame(self, corner_radius=15)
        file_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(file_frame, text="Choose Excel File:", font=("Poppins", 14)).pack(pady=5)

        choose_button = ctk.CTkButton(file_frame, text="Browse File", 
                                      command=self.choose_file, width=200)
        choose_button.pack(pady=5)

        self.file_label = ctk.CTkLabel(file_frame, text="No file selected", 
                                       text_color="gray80", font=("Poppins", 12))
        self.file_label.pack(pady=5)

        # ---------- MESSAGE BOX ----------
        msg_frame = ctk.CTkFrame(self, corner_radius=15)
        msg_frame.pack(pady=10, padx=20, fill="both")

        ctk.CTkLabel(msg_frame, text="Message Template:", font=("Poppins", 14)).pack(pady=5)

        self.message_box = ctk.CTkTextbox(msg_frame, height=120, width=520, corner_radius=12)
        self.message_box.insert("1.0", "Assalamualaikum {name},\nThis is an automated message.")
        self.message_box.pack(pady=10)

        # ---------- SEND BUTTON ----------
        send_button = ctk.CTkButton(self, text="Start Sending", 
                                    font=("Poppins", 16, "bold"),
                                    command=self.start_sending, height=45, corner_radius=10)
        send_button.pack(pady=10)

        # ---------- VIEW CONTACTS BUTTON ----------
        view_button = ctk.CTkButton(self, text="View Contacts & Status",
                                     font=("Poppins", 14, "bold"),
                                     command=self.show_contacts, height=40, corner_radius=10)
        view_button.pack(pady=5)

        # ---------- LOG AREA ----------
        log_frame = ctk.CTkFrame(self, corner_radius=15)
        log_frame.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkLabel(log_frame, text="Log Output:", font=("Poppins", 14)).pack(pady=5)

        self.log_window = ctk.CTkTextbox(log_frame, height=200, width=700)
        self.log_window.pack(padx=10, pady=10)

    # -------------------------------------
    #         Choose Excel File
    # -------------------------------------
    def choose_file(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.excel_path:
            self.file_label.configure(text=self.excel_path, text_color="white")

    # -------------------------------------
    #             Logging
    # -------------------------------------
    def log(self, text):
        self.log_window.insert("end", text + "\n")
        self.log_window.see("end")
        self.update()

    # -------------------------------------
    #          Show Contacts & Status
    # -------------------------------------
    def show_contacts(self):
        contacts_window = ctk.CTkToplevel(self)
        contacts_window.title("Contacts & Status")
        contacts_window.geometry("600x400")

        frame = ctk.CTkFrame(contacts_window, corner_radius=15)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        textbox = ctk.CTkTextbox(frame, width=580, height=350)
        textbox.pack(padx=10, pady=10, fill="both", expand=True)

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT phone, name, status, timestamp FROM contacts")
        rows = c.fetchall()
        conn.close()

        textbox.insert("0.0", f"{'Phone':<15}{'Name':<20}{'Status':<10}{'Timestamp':<20}\n")
        textbox.insert("0.0", "-"*70 + "\n")
        for row in rows:
            phone, name, status, timestamp = row
            timestamp = timestamp if timestamp else "-"
            textbox.insert("end", f"{phone:<15}{name:<20}{status:<10}{timestamp:<20}\n")

        textbox.configure(state="disabled")  # make read-only

    # -------------------------------------
    #          Main Sending Logic
    # -------------------------------------
    def start_sending(self):
        if not self.excel_path:
            self.log("❌ Please select an Excel file first.")
            return

        message_template = self.message_box.get("1.0", "end").strip()

        try:
            df = pd.read_excel(self.excel_path)
        except:
            self.log("❌ Could not read Excel file.")
            return

        if "number" not in df.columns:
            self.log("❌ Excel missing 'number' column.")
            return

        df["number"] = df["number"].apply(clean_phone)
        df = df[df["number"] != ""].reset_index(drop=True)

        if df.empty:
            self.log("❌ No valid phone numbers.")
            return

        # ---------- Add contacts to database ----------
        for idx, row in df.iterrows():
            phone = clean_phone(row["number"])
            name = row.get("name", "")
            add_contact(phone, name)

        self.log("🚀 Starting Chrome…")

        # ---------- Selenium Setup ----------
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(r"--user-data-dir=C:\chrome-whatsapp-profile")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=options)

        driver.get("https://web.whatsapp.com")
        self.log("📱 Login to WhatsApp Web (scan QR if needed)…")

        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
            )
        except:
            self.log("❌ Login failed (timeout).")
            return

        self.log("✅ Logged in! Sending messages…")

        # ---------- Sending Loop with DB update ----------
        for i, row in df.iterrows():
            phone = row["number"]
            name = row.get("name", "")

            try:
                msg = message_template.format(name=name)
            except:
                msg = message_template

            encoded = urllib.parse.quote(msg)

            self.log(f"➡ Sending to {phone} ({i+1}/{len(df)})…")
            
            success = send_message_to_number(driver, phone, encoded, self.log)
            if success:
                update_status(phone, "Sent")
            else:
                update_status(phone, "Failed")
            
            time.sleep(4)

        self.log("🎉 All messages sent!")
        driver.quit()

# ---------------- RUN APP ----------------
if __name__ == "__main__":
    init_db()
    app = WhatsAppModernApp()
    app.mainloop()