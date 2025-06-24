import os
import sys
import time
import subprocess
import pandas as pd
import pyautogui
import pyperclip
import pygetwindow as gw
from tkinter import messagebox, Tk
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tqdm import tqdm
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font
import re


class HaryanaEBillingemptytbl:
    def __init__(self, website_url):
        self.website_url = website_url  # Use the parameter!
        self.driver = None
        self.wait = None
        self.all_data = []
        self.wait_time = 15
        self.max_retries = 3
        self._chrome_was_launched = False

    def _show_console(self):
        """Ensure console window is visible"""
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 1)

    def _check_prerequisites(self):
        """Verify all requirements are met"""
        self._show_console()

        print("="*50)
        print("CHECKING PREREQUISITES")
        print("="*50)
        
        # 1. Verify packages
        try:
            import selenium, pyautogui, pyperclip, pygetwindow
            print("✅ Required packages are installed")
        except ImportError as e:
            print(f"❌ Missing package: {str(e)}")
            print("Please run: pip install selenium pyautogui pyperclip pygetwindow")
            messagebox.showerror("Missing Packages", f"Please install:\nselenium\npyautogui\npyperclip\npygetwindow")
            sys.exit(1)
            
        # 2. Check Chrome debug status
        if not self._is_chrome_running_with_debug():
            print("⚠️ Chrome not running with debugging port")
            if not self._launch_chrome_with_debug():
                messagebox.showerror("Chrome Error", "Failed to launch Chrome with debugging")
                sys.exit(1)
            self._chrome_was_launched = True
            
        return True

    def _is_chrome_running_with_debug(self):
        """Check if Chrome is already running with debug port"""
        try:
            result = subprocess.run(
                ['netstat', '-ano'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return ":9222" in result.stdout
        except:
            return False

    def _launch_chrome_with_debug(self):
        """Launch Chrome with remote debugging"""
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]
        for path in chrome_paths:
            if os.path.exists(path):
                print(f"🚀 Launching Chrome with debugging: {path}")
                try:
                    subprocess.Popen([
                        path,
                        "--remote-debugging-port=9222",
                        "--user-data-dir=C:\\Temp\\ChromeDebugProfile",
                        self.website_url
                    ], creationflags=subprocess.CREATE_NO_WINDOW)
                    time.sleep(1)
                    sys.exit(0)  # Exit after launching Chrome
                except Exception as e:
                    print(f"Failed to launch Chrome: {str(e)}")
                    return False
                    
        print("❌ Chrome not found in standard locations")
        print("Please manually start Chrome with:")
        print("chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\\Temp\\ChromeDebugProfile")
        return False

    def connect_to_browser(self):
        """Connect to existing Chrome browser instance"""
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        print("="*50)
        print("BROWSER CONNECTION")
        print("="*50)
        
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, self.wait_time)
            
            if self._chrome_was_launched:
                print("\n⚠️ NEW CHROME SESSION DETECTED")
                print("Please complete login to Haryana e-Billing in the Chrome window")
                print("After login, rerun script to continue...")
                sys.exit(0)  # Exit gracefully, no error
            else:
                print("✅ Reconnected to existing Chrome session")
                
            print("="*50)
            print("STARTING AUTOMATION")
            print("="*50)
            return True
            
        except Exception as e:
            print(f"❌ Browser connection failed: {str(e)}")
            return False

    def ensure_window_visible(self):
        """Bring Chrome window to foreground if minimized"""
        try:
            chrome_windows = gw.getWindowsWithTitle("Chrome")
            if chrome_windows:
                chrome_window = chrome_windows[0]
                if chrome_window.isMinimized:
                    chrome_window.restore()
                chrome_window.activate()
                time.sleep(0.5)
                return True
        except Exception as e:
            print(f"Window management warning: {str(e)}")
        return False


    def get_dropdown_options(self, dropdown_id):
        """Get all non-empty, non-placeholder options from a dropdown"""
        try:
            dropdown = Select(self.wait.until(
                EC.presence_of_element_located((By.ID, dropdown_id))
            ))
            return [
                {'value': opt.get_attribute("value"), 'text': opt.text.strip()}
                for opt in dropdown.options
                if opt.get_attribute("value")
                and opt.text.strip()
                and opt.text.strip().lower() not in ["select one", "select", "choose", "--select--"]
            ]
        except Exception as e:
            print(f"⚠ Error reading {dropdown_id} options: {str(e)}")
            return []

    def safe_select_dropdown(self, dropdown_id, value):
        """Select dropdown option with retries and proper waiting"""
        for attempt in range(self.max_retries):
            try:
                dropdown = Select(self.wait.until(
                    EC.element_to_be_clickable((By.ID, dropdown_id))
                ))
                dropdown.select_by_value(value)
                
                # Wait for the dropdown value to update
                self.wait.until(lambda d: dropdown.first_selected_option.get_attribute("value") == value)
                # Wait for dependent element to update (if needed)
                self.wait.until(EC.presence_of_element_located((By.ID, "lbltobeexecuted")))
                return True
                
            except Exception as e:
                print(f"⚠ Attempt {attempt+1} failed for {dropdown_id}={value}: {str(e)}")
                time.sleep(0.5)  # Reduced sleep
        
        print(f"✖ Failed to select {dropdown_id}={value} after {self.max_retries} attempts")
        return False



    def delete_table(self):
        """Delete all rows in the current table in a safe order (positive, negative, positive, ...) to avoid exceeding limits."""
        try:
            while True:
                try:
                    table = self.wait.until(
                        EC.presence_of_element_located((By.ID, "GV_Add_to_List_POP"))
                    )
                except TimeoutException:
                    print("Table not found. Assuming all rows deleted.")
                    break

                rows = table.find_elements(By.XPATH, ".//tr[not(contains(@class, 'header-row')) and not(contains(@class, 'total-row'))]")
                if not rows:
                    print("No more rows to delete.")
                    break

                # Parse rows into positive and negative qty
                pos_rows = []
                neg_rows = []
                for row in rows:
                    try:
                        qty_cell = row.find_elements(By.TAG_NAME, "td")[10]
                        qty_val = float(qty_cell.text.strip())
                        if qty_val >= 0:
                            pos_rows.append(row)
                        else:
                            neg_rows.append(row)
                    except Exception:
                        continue

                # Alternate deletion: positive, negative, positive, ...
                delete_sequence = []
                i, j = 0, 0
                while i < len(pos_rows) or j < len(neg_rows):
                    if i < len(pos_rows):
                        delete_sequence.append(pos_rows[i])
                        i += 1
                    if j < len(neg_rows):
                        delete_sequence.append(neg_rows[j])
                        j += 1

                if not delete_sequence:
                    print("No deletable rows found in table.")
                    break

                # Delete the first in sequence
                row = delete_sequence[0]
                try:
                    delete_btns = row.find_elements(By.XPATH, ".//a[contains(@id, 'lnkDelete_POP')]")
                    if not delete_btns:
                        print("No delete button found in this row. Table might be empty now.")
                        break
                    delete_btn = delete_btns[0]
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", delete_btn)
                    delete_btn.click()
                    # Accept browser confirmation alert
                    try:
                        WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except TimeoutException:
                        print("No browser alert appeared.")
                    # Wait for the success message and click OK
                    try:
                        ok_btn = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.confirm"))
                        )
                        ok_btn.click()
                    except TimeoutException:
                        print("No success OK button appeared.")
                    # Wait for table to reload
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠ Error deleting row: {str(e)}")
                    break

            print("All rows deleted from table.")
        except Exception as e:
            print(f"⚠ Error deleting table rows: {str(e)}")

    def delete_all_combinations(self):
        """Delete table rows for all dropdown combinations."""
        print("\nStarting deletion process for all combinations...")

        main_heads = [
            mh for mh in self.get_dropdown_options("ddlcomp")
            if mh['text'].strip() and mh['text'].strip().lower() != "select one"
        ]
        print(f"Found {len(main_heads)} Main Head options (excluding empty/'Select One')")

        for main_head in tqdm(main_heads, desc="Main Heads"):
            print(f"\nProcessing Main Head: {main_head['text']}")
            if not self.safe_select_dropdown("ddlcomp", main_head['value']):
                continue

            sub_heads = [
                sh for sh in self.get_dropdown_options("ddlsubhead")
                if sh['text'].strip() and sh['text'].strip().lower() != "select one"
            ]
            print(f"├─ Found {len(sub_heads)} Sub Head options (excluding empty/'Select One')")

            for sub_head in sub_heads:
                if not self.safe_select_dropdown("ddlsubhead", sub_head['value']):
                    continue

                items = [
                    it for it in self.get_dropdown_options("ddlitemnumber")
                    if it['text'].strip() and it['text'].strip().lower() != "select one"
                ]

                for item in items:
                    if not self.safe_select_dropdown("ddlitemnumber", item['value']):
                        continue

                    # Delete all rows in the table for this combination
                    self.delete_table()

        print("\n✔ Completed! All table rows deleted for all combinations.")

    def close(self):
        if self.driver:
            self.driver.quit()

def ask_delete_mode():
    """Premium dialog for delete mode selection with watermark and centered window."""
    root = tk.Tk()
    root.withdraw()

    result = {"choice": None}

    def set_result(val):
        result["choice"] = val
        win.destroy()
        root.quit()

    def on_close():
        result["choice"] = "cancel"
        win.destroy()
        root.quit()

    win = tk.Toplevel(root)
    win.title("Delete Table Data")
    win.resizable(False, False)

    # Calculate center position
    win.update_idletasks()
    width, height = 400, 250
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")
    win.attributes('-topmost', True)

    # Premium style
    frame = tk.Frame(win, bg="#f5f7fa", bd=2, relief="groove")
    frame.place(relwidth=1, relheight=1)

    tk.Label(frame, text="What do you want to delete?", 
             font=("Segoe UI Semibold", 13), bg="#f5f7fa", fg="#222").pack(pady=(18, 8))

    btn_style = {"font": ("Segoe UI", 11), "bg": "#0078D7", "fg": "white", "activebackground": "#005A9E", "activeforeground": "white", "bd": 0, "relief": "flat", "cursor": "hand2"}

    # Helper for hover effect (color + size)
    def on_enter(e, btn, color):
        btn['bg'] = color
        btn['font'] = ("Segoe UI", 12, "bold")  # Slightly larger and bold
        btn['width'] = 30                       # Slightly wider

    def on_leave(e, btn, color):
        btn['bg'] = color
        btn['font'] = ("Segoe UI", 11)
        btn['width'] = 28

    # Delete ALL Combinations button
    btn_all = tk.Button(frame, text="Delete ALL Combinations", width=28,
                        command=lambda: set_result("all"), **btn_style)
    btn_all.pack(pady=4)
    btn_all.bind("<Enter>", lambda e: on_enter(e, btn_all, "#005A9E"))
    btn_all.bind("<Leave>", lambda e: on_leave(e, btn_all, "#0078D7"))

    # Delete CURRENT Table Only button
    btn_current = tk.Button(frame, text="Delete CURRENT Table Only", width=28,
                            command=lambda: set_result("current"), **btn_style)
    btn_current.pack(pady=4)
    btn_current.bind("<Enter>", lambda e: on_enter(e, btn_current, "#005A9E"))
    btn_current.bind("<Leave>", lambda e: on_leave(e, btn_current, "#0078D7"))

    # Cancel button at the bottom
    btn_cancel = tk.Button(frame, text="Cancel", width=28,
                           command=lambda: set_result("cancel"),
                           font=("Segoe UI", 11), bg="#e81123", fg="white",
                           activebackground="#a80000", activeforeground="white", bd=0, relief="flat", cursor="hand2")
    btn_cancel.pack(side="bottom", pady=14)
    btn_cancel.bind("<Enter>", lambda e: on_enter(e, btn_cancel, "#a80000"))
    btn_cancel.bind("<Leave>", lambda e: on_leave(e, btn_cancel, "#e81123"))

    # Watermark (above the cancel button)
    tk.Label(frame, text="Developed by MRGARGSIR", font=("Segoe UI", 9, "italic"),
             fg="#b0b0b0", anchor="se").place(relx=1.0, rely=1.0, x=-100, y=-60, anchor="se")

    win.protocol("WM_DELETE_WINDOW", on_close)
    win.grab_set()
    root.mainloop()
    root.destroy()
    return result["choice"]

    # In your main function:
def main():
    print("Haryana E-Billing Data emptytbl")
    print("=" * 50)
    
   
    
    emptytbl = HaryanaEBillingemptytbl(
        "https://works.haryana.gov.in/E-Billing/Est_Add_Items_emb.aspx#"
    )
    
    try:
        print("[1/5] Checking prerequisites...")
        if not emptytbl._check_prerequisites():
            return
        
        print("[2/5] Connecting to browser...")
        if not emptytbl.connect_to_browser():
            return
            
        print("[3/5] Ensuring window visibility...")
        emptytbl.ensure_window_visible()

        print("[4/5] Asking for deletion mode...")
        mode = ask_delete_mode()
        print(f"User selected: {mode}")
        
        if mode == "cancel":
            print("Operation cancelled by user.")
            return
            
        if mode == "all":
            print("[5/5] Deleting ALL combinations...")
            emptytbl.delete_all_combinations()
        elif mode == "current":
            print("[5/5] Deleting CURRENT table...")
            emptytbl.delete_table()
            
        print("\nOperation completed successfully!")
        
    except Exception as e:
        print(f"\n✖ Critical error: {str(e)}")
        import traceback
        traceback.print_exc()
        
    finally:
        print("Cleaning up...")
        emptytbl.close()
        try:
            tk._default_root.destroy()
        except:
            pass
if __name__ == "__main__":
    main()
