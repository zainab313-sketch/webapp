import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import time
import urllib.parse
import threading

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import sqlite3
import datetime

# ===== Glass Theme Colors =====
GLASS_BG = "#96608c"
GLASS_CARD = "#a690b4"
GLASS_BORDER = "#4b3d57"

PRIMARY = "#4C5E74"          # WhatsApp green
PRIMARY_HOVER = "#519be1"

TEXT_MAIN = "#1e1e1e"
TEXT_MUTED = "#6b7280"

SUCCESS = "#4CAF50"
ERROR = "#E53935"

# ---------------- Driver ----------------
def init_driver():
    options = webdriver.ChromeOptions()
    # options.add_argument(r"--user-data-dir=C:\chrome-whatsapp-profile")
    options.add_argument(r"--user-data-dir=C:\chrome-whatsapp-profile-selenium")

    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    

    try:
       driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
       driver.get("https://web.whatsapp.com")
       return driver
    except Exception as e:
        print("[ERROR] Failed to start Chrome:", e)
        return None

def wait_for_login(driver, timeout=60):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
        )
        return True
    except:
        return False

# ---------------- Database ----------------
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

# ---------------- Phone cleaning ----------------

def clean_phone(phone) -> str:
    if pd.isna(phone):
        return ""

    # Convert Excel floats safely
    phone = str(phone).strip()

    # Remove trailing .0 if exists
    if phone.endswith(".0"):
        phone = phone[:-2]

    # Keep only digits
    phone = "".join(ch for ch in phone if ch.isdigit())

    return phone




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

# ---------------- WhatsApp ----------------
def send_message_to_contacts(driver, phone, encoded_message):
    url = f"https://web.whatsapp.com/send?phone={phone}&app_absent=0"
    driver.get(url)

    try:
        # Wait for message box
        chat_box = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            )
        )

        time.sleep(1)

        # Decode message back to text
        # message = urllib.parse.unquote(encoded_message)

        # Clear + type message like a human
        chat_box.click()
        chat_box.send_keys(encoded_message)
        time.sleep(1)

        # Send message
        chat_box.send_keys(Keys.ENTER)
        time.sleep(2)

        return True

    except Exception as e:
        print(f"[!] Failed to send to {phone}: {e}")
        return False


# ---------------- GUI ----------------
class WhatsAppModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("WhatsApp Bulk Messenger — Modern Edition")
        self.geometry("780x700")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")  # base only (required)

        self.configure(fg_color=GLASS_BORDER)
       
        self.excel_path = ""

        # Header
        header = ctk.CTkLabel(
        self,
        text="WhatsApp Bulk Messaging App",
        font=("Poppins", 28, "bold"),
        text_color=TEXT_MAIN
    )
        header.pack(pady=12)
        

        # File select
        file_frame = ctk.CTkFrame(
        self,
        fg_color=GLASS_CARD,
        border_color=GLASS_BORDER,
        border_width=1,
        corner_radius=18
    )
        file_frame.pack(pady=10, padx=20, fill="x")
       
        ctk.CTkLabel(file_frame, text="Choose Excel File:", font=("Poppins", 14)).pack(pady=5)
        choose_button = ctk.CTkButton(
            file_frame,
            text="Browse File",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            corner_radius=16,
            width=200,
            command=self.choose_file
        )
        choose_button.pack(pady=5)
        
        self.file_label = ctk.CTkLabel(file_frame, text="No file selected", text_color="gray80", font=("Poppins", 12))
        self.file_label.pack(pady=5)

        # Message box
        msg_frame = ctk.CTkFrame(
            self,
            fg_color=GLASS_CARD,
            border_color=GLASS_BORDER,
            border_width=1,
            corner_radius=18
        )
        msg_frame.pack(pady=10, padx=20, fill="both")
     
        ctk.CTkLabel(msg_frame, text="Message Template:", font=("Poppins", 14)).pack(pady=5)
        self.message_box = ctk.CTkTextbox(
                msg_frame,
                height=80,
                width=300,
                corner_radius=16,
                fg_color="#f9fafb",
                text_color=TEXT_MAIN,
                border_color=GLASS_BORDER,
                border_width=1
            )
        self.message_box.insert("1.0", "Assalamualaikum,\nThis is an automated message.")
        self.message_box.pack(pady=10)

         # -------- Status Filter --------
        filter_frame = ctk.CTkFrame(
            self,
            fg_color=GLASS_CARD,
            border_color=GLASS_BORDER,
            border_width=1,
            corner_radius=18
        )
        filter_frame.pack(pady=10, padx=20, fill="x")

        self.only_applied_var = ctk.BooleanVar(value=True)

        self.only_applied_checkbox = ctk.CTkCheckBox(
            filter_frame,
            text="Send messages only to contacts with status = 'Applied'",
            variable=self.only_applied_var,
            text_color=TEXT_MAIN,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER
        )

        self.only_applied_checkbox.pack(pady=10, padx=10, anchor="w")


        # Buttons
        ##-------Send Btn------##
        send_button = ctk.CTkButton(
            self,
            text="Start Sending",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=("Poppins", 16, "bold"),
            height=42,
            corner_radius=18,
            command=self.start_sending_thread
        )
        send_button.pack(pady=10)
        ##-------View Btn------##
        view_button = ctk.CTkButton(
            self,
            text="View Contacts & Status",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=("Poppins", 16, "bold"),
            height=42,
            corner_radius=18,
            command=self.show_contacts
        )
        view_button.pack(pady=10)
       
        log_frame = ctk.CTkFrame(
            self,
            fg_color=GLASS_CARD,
            border_color=GLASS_BORDER,
            border_width=1,
            corner_radius=18
        )
        log_frame.pack(pady=8, padx=20, fill="x")
        ctk.CTkLabel(log_frame, text="Log Output:", font=("Poppins", 14)).pack(pady=1)
        self.log_window = ctk.CTkTextbox(
            log_frame,
            height=600,
            width=700,
            corner_radius=13,
            fg_color="#f9fafb",
            text_color=TEXT_MAIN,
            border_color=GLASS_BORDER,
            border_width=1
        )
        self.log_window.pack(padx=10, pady=1)

    def find_column(self, df, possible_names):
        """
        Finds a column in df matching any name in possible_names (case-insensitive)
        """
        lower_cols = {col.lower(): col for col in df.columns}
        for name in possible_names:
            if name.lower() in lower_cols:
                return lower_cols[name.lower()]
        return None
    

    # ---------------- GUI Methods ----------------
    def choose_file(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.excel_path:
            self.file_label.configure(text=self.excel_path, text_color="white")

    def log(self, text):
        self.log_window.insert("end", text + "\n")
        self.log_window.see("end")
        self.update()

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
        textbox.configure(state="disabled")

    # ---------------- Thread ----------------
    def start_sending_thread(self):
        threading.Thread(target=self.start_sending, daemon=True).start()

    # ---------------- Sending ----------------
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

        contacts_col = self.find_column(df, ["contacts", "contact", "phone", "number", "mobile", "whatsapp"])
        if not contacts_col:
            self.log("❌ No contacts column found.")
            return
        name_col = self.find_column(df, ["name", "full name", "candidate name"])

        # Apply status filter if needed
        status_col = self.find_column(df, ["status", "application status", "state"])
        if status_col and self.only_applied_var.get():
            df = df[df[status_col].astype(str).str.strip().str.lower() == "applied"].reset_index(drop=True)

        # Clean and normalize contacts
        df[contacts_col] = df[contacts_col].apply(clean_phone).apply(normalize_contacts)
        df = df[df[contacts_col].str.len() == 12].reset_index(drop=True)
        if df.empty:
            self.log("❌ No valid phone numbers after cleaning.")
            return

        print("📋 Numbers after cleaning:")
        print(df[contacts_col].tolist())

        
        if "WhatsApp Status" not in df.columns:
            df["WhatsApp Status"] = ""

        # Add contacts to DB
        for _, row in df.iterrows():
            phone = row[contacts_col]
            name = row[name_col] if name_col else ""
            add_contact(phone, name)

        # ---------------- Chrome Login ----------------
        self.log("🚀 Launching Chrome for login…")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument(f"--user-data-dir=C:\\chrome-whatsapp-profile")  # persistent login
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        except Exception as e:
            self.log(f"❌ Could not start Chrome: {e}")
            return

        driver.get("https://web.whatsapp.com")
        self.log("📱 Please scan QR code if needed…")
        if not wait_for_login(driver, timeout=120):
            self.log("❌ Login failed or timed out.")
            driver.quit()
            return
        self.log("✅ Logged into WhatsApp Web")

        # ---------------- Sending Loop ----------------
        for idx, row in df.iterrows():
            phone = row[contacts_col]
            name = row[name_col] if name_col else ""

             # Skip already sent
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT status FROM contacts WHERE phone=?", (phone,))
            result = c.fetchone()
            conn.close()
            if result and result[0] == "Sent":
                self.log(f"⚠ Skipping {phone} — already sent.")
                continue
            try:
                msg = message_template.format(name=name)
            except:
                msg = message_template
            # encoded_msg = msg 
            encoded_msg = urllib.parse.quote(msg)
            self.log(f"➡ Sending to {phone} ({idx+1}/{len(df)})…")
            success = send_message_to_contacts(driver, phone, encoded_msg)
            if success:
                update_status(phone, "Sent")
                df.loc[idx, "WhatsApp Status"] = ""  # leave blank for valid
            else:
                update_status(phone, "Failed")
                df.loc[idx, "WhatsApp Status"] = "Invalid"  # mark invalid

            # Save Excel after each row (optional, safe)
            if idx % 10 == 0:
                df.to_excel(self.excel_path, index=False)

            df.to_excel(self.excel_path, index=False)
            time.sleep(4)
        self.log("🎉 All messages processed! Closing browser in 5 seconds…")
        time.sleep(5)
        driver.quit()
#     def start_sending(self):
#         if not self.excel_path:
#             self.log("❌ Please select an Excel file first.")
#             return

#         message_template = self.message_box.get("1.0", "end").strip()

#         try:
#             df = pd.read_excel(self.excel_path)
#         except:
#             self.log("❌ Could not read Excel file.")
#             return
        
#     # Detect contacts column automatically
#         contacts_col = self.find_column(
#             df,
#             ["contacts", "contact", "phone", "number", "mobile", "whatsapp"]
#         )

#         if not contacts_col:
#             self.log("❌ No contacts column found (Contacts / Phone / Number / Mobile).")
#             return
        
#         name_col = self.find_column(df, ["name", "full name", "candidate name"])


# # Apply status filter ONLY if column exists
#         status_col = self.find_column(df, ["status", "application status", "state"])

#         if status_col :
#             if self.only_applied_var.get():
#                 self.log("ℹ Status column found — sending only to 'Applied' contacts.")
#                 df = df[
#                     df[status_col]
#                     .astype(str)
#                     .str.strip()
#                     .str.lower() == "applied"
#                 ].reset_index(drop=True)
#             else:
#                self.log("ℹ Status column found — checkbox OFF, sending to ALL contacts.")
#         else:
#            self.log("ℹ No status column found — sending to ALL contacts.")

#         ##=====cleaning contacts=====##

#         df[contacts_col] = (
#             df[contacts_col]
#             .apply(clean_phone)
#             .apply(normalize_contacts)
#         )

#         print("📋 Numbers after cleaning:")
#         print(df[contacts_col].head(10).tolist())


#         df = df[df[contacts_col].str.len() == 12].reset_index(drop=True)


#         if df.empty:
#             self.log("❌ No valid phone numbers after cleaning.")
#             return

        
#         if "WhatsApp Status" not in df.columns:
#             df["WhatsApp Status"] = ""

#         # Add contacts to DB
#         for _, row in df.iterrows():
#             phone = row[contacts_col]
#             name = row[name_col] if name_col else ""
#             add_contact(phone, name)

#         # ---------------- Chrome Login ----------------
#         self.log("🚀 Launching Chrome for login…")
#         driver = init_driver()
#         if not driver:
#             self.log("❌ Could not start Chrome.")
#             return

#         self.log("📱 Please scan QR if needed…")
#         if not wait_for_login(driver, timeout=120):
#             self.log("❌ Login failed or timeout.")
#             driver.quit()
#             return
#         self.log("✅ Login successful!")


#         # ---------------- Sending Loop 1----------------
#         for i, row in df.iterrows():
#             phone = row[contacts_col]
#             name = row[name_col] if name_col else ""

#             # Skip already sent
#             conn = sqlite3.connect(DB_FILE)
#             c = conn.cursor()
#             c.execute("SELECT status FROM contacts WHERE phone=?", (phone,))
#             result = c.fetchone()
#             conn.close()
#             if result and result[0] == "Sent":
#                 self.log(f"⚠ Skipping {phone} — already sent.")
#                 continue

#             try:
#                 msg = message_template.format(name=name)
#             except:
#                 msg = message_template

#             encoded = urllib.parse.quote(msg)
#             self.log(f"➡ Sending to {phone} ({i+1}/{len(df)})…")
#             success = send_message_to_contacts(driver, phone, encoded)

#             if success:
#                 update_status(phone, "Sent")
#                 df.loc[i, "WhatsApp Status"] = ""  # leave blank for valid
#             else:
#                 update_status(phone, "Failed")
#                 df.loc[i, "WhatsApp Status"] = "Invalid"  # mark invalid

#             # Save Excel after each row (optional, safe)
#             if i % 10 == 0:
#                 df.to_excel(self.excel_path, index=False)

#             df.to_excel(self.excel_path, index=False)
#             time.sleep(4)

#         self.log("🎉 All messages sent!")
#         driver.quit()


# ---------------- RUN ----------------
if __name__ == "__main__":
    init_db()
    app = WhatsAppModernApp()
    app.mainloop()