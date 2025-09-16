import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.parse

# ------------------------
# CONFIGURATION
# ------------------------
CHROMEDRIVER_PATH = r"C:\Users\ahtis\OneDrive\Documents\WebDriver\chromedriver-win64\chromedriver.exe"
MESSAGE_TEMPLATE = "Hello {name}, this is a test message sent through WhatsApp automation!"

# ------------------------
# GUI: Excel File Upload
# ------------------------
def browse_file():
    global EXCEL_FILE_PATH
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        EXCEL_FILE_PATH = file_path
        file_label.config(text=f"Selected: {file_path}")
        send_button.config(state="normal")

def start_script():
    root.destroy()  # Close the GUI and continue to automation

# GUI layout
EXCEL_FILE_PATH = None
root = tk.Tk()
root.title("WhatsApp Automation (Send by Number)")
root.geometry("480x250")

tk.Label(root, text="Step 1: Upload Excel with 'Name' and 'Phone Number' columns").pack(pady=10)
tk.Button(root, text="Browse Excel", command=browse_file).pack(pady=5)
file_label = tk.Label(root, text="No file selected")
file_label.pack(pady=5)

tk.Label(root, text="Step 2: Log in to WhatsApp & Send Messages").pack(pady=20)
send_button = tk.Button(root, text="Start WhatsApp Messaging", command=start_script, state="disabled", bg="green", fg="white")
send_button.pack(pady=10)

root.mainloop()

# ------------------------
# AUTOMATION LOGIC
# ------------------------
if EXCEL_FILE_PATH:
    # Load Excel
    df = pd.read_excel(EXCEL_FILE_PATH, dtype=str)
    df.fillna('', inplace=True)
    contacts = df.to_dict(orient='records')

    # Set up browser
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get("https://web.whatsapp.com")
    input("📱 Please scan QR code in the browser and press ENTER here to continue...")

    total_sent = 0
    failed_numbers = []

    for contact in contacts:
        name = contact.get("Name", "").strip()
        number = contact.get("Phone Number", "").strip().replace(" ", "").replace("+", "")

        if not number.isdigit():
            print(f"⚠️ Skipping invalid number: {number}")
            failed_numbers.append(number)
            continue

        # Prepare message
        message = MESSAGE_TEMPLATE.format(name=name)

        try:
            # Click search bar and enter number
            search_box = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
            )
            search_box.clear()
            search_box.click()
            time.sleep(0.5)
            search_box.send_keys(number)
            time.sleep(2)
            search_box.send_keys(Keys.ENTER)

            # Wait for message box and send message
            message_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            message_box.send_keys(message)
            message_box.send_keys(Keys.ENTER)

            print(f"✅ Sent to {number} ({name})")
            total_sent += 1
            time.sleep(2)

        except Exception as e:
            print(f"❌ This number does not have WhatsApp: {number} ({name})")
            failed_numbers.append(number)
            time.sleep(2)
            continue

    driver.quit()

    summary = f"✅ Messages sent: {total_sent}\n❌ Failed: {len(failed_numbers)}"
    if failed_numbers:
        summary += "\n\nFailed numbers:\n" + "\n".join(failed_numbers)

    messagebox.showinfo("Done", summary)
else:
    messagebox.showwarning("No Excel", "You must select an Excel file to start.")
