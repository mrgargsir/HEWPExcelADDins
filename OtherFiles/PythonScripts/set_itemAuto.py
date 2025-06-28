import os
import sys
import ctypes
import time
import subprocess
import importlib.util

import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import traceback

import pyperclip
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException, TimeoutException
from selenium.common.exceptions import NoSuchElementException
import time
from selenium.common.exceptions import StaleElementReferenceException

class HEWPSetItem:
    def __init__(self):
        print("[INIT] Initializing HEWPSetItem...")
        self.notification_root = None
        self._chrome_was_launched = False
        self.driver = None
        self.wait = None

        self._hide_console()

        if self._check_prerequisites():
            print("[INIT] Prerequisites OK. Connecting to browser...")
            self.connect_to_browser()
            self._hide_console()

    def _show_console(self):
        if sys.platform == 'win32':
            print("[CONSOLE] Showing console window.")
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 1)

    def _hide_console(self):
        if sys.platform == 'win32':
            print("[CONSOLE] Hiding console window.")
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    def _check_prerequisites(self):
        """Verify all requirements are met"""
        print("[CHECK] Checking prerequisites...")
        print("="*50)
        print("CHECKING PREREQUISITES")
        print("="*50)

        required_modules = [
            "selenium", "pyperclip", "pygetwindow", "tkinter"
        ]
        missing = []
        for mod in required_modules:
            if importlib.util.find_spec(mod) is None:
                missing.append(mod)
        if missing:
            print(f"❌ Missing packages: {', '.join(missing)}")
            print(f"Please run: pip install {' '.join(missing)}")
            messagebox.showerror("Missing Packages", f"Please install:\n" + "\n".join(missing))
            sys.exit(1)
        else:
            print("✅ All required Python packages are installed.")

        # Check Chrome debug status
        if not self._is_chrome_running_with_debug():
            print("⚠️ Chrome not running with debugging port")
            if not self._launch_chrome_with_debug():
                messagebox.showerror("Chrome Error", "Failed to launch Chrome with debugging")
                sys.exit(1)
            self._chrome_was_launched = True
        else:
            print("✅ Chrome is running with debugging port.")
        return True

    def _is_chrome_running_with_debug(self):
        """Check if Chrome is running with debug port and is a chrome.exe process.
        If port is open but not by chrome.exe, kill the process and exit."""
        print("[CHECK] Checking if Chrome is running with debug port...")
        try:
            result = subprocess.run(
                ['netstat', '-ano'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            lines = result.stdout.splitlines()
            for line in lines:
                if ":9222" in line and "LISTENING" in line:
                    parts = line.split()
                    pid = parts[-1]
                    # Now check if this PID is a Chrome process
                    tasklist = subprocess.run(
                        ['tasklist', '/FI', f'PID eq {pid}'],
                        capture_output=True,
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if "chrome.exe" in tasklist.stdout.lower():
                        print(f"[CHECK] Chrome debug port found and process is chrome.exe (PID {pid})")
                        return True
                    else:
                        print(f"[CHECK] Port 9222 is open but not used by chrome.exe (PID {pid})")
                        # Try to kill the process using the port
                        try:
                            subprocess.run(['taskkill', '/PID', pid, '/F'], creationflags=subprocess.CREATE_NO_WINDOW)
                            print(f"[CHECK] Killed process PID {pid} using port 9222.")
                        except Exception as kill_err:
                            print(f"[CHECK] Failed to kill process PID {pid}: {kill_err}")
                        print("[CHECK] Exiting script due to invalid debug port usage.")
                        import sys
                        sys.exit(1)
            print("[CHECK] Chrome debug port not found or not a Chrome process.")
            return False
        except Exception as e:
            print(f"[CHECK] Error checking Chrome debug port: {e}")
            return False

    def _launch_chrome_with_debug(self):
        print("[LAUNCH] Attempting to launch Chrome with debugging...")
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
                        "https://works.haryana.gov.in/HEWP_Login/login.aspx"
                    ], creationflags=subprocess.CREATE_NO_WINDOW)
                    time.sleep(0.2)
                    print("[LAUNCH] Chrome launched. Please login and rerun script.")
                    sys.exit(0)
                except Exception as e:
                    print(f"Failed to launch Chrome: {str(e)}")
                    return False

        print("❌ Chrome not found in standard locations")
        print("Please manually start Chrome with:")
        print("chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\\Temp\\ChromeDebugProfile")
        return False

    def connect_to_browser(self):
        """Connect to existing Chrome browser instance"""
        print("[BROWSER] Connecting to Chrome browser...")
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        print("="*50)
        print("BROWSER CONNECTION")
        print("="*50)
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            if self._chrome_was_launched:
                print("\n⚠️ NEW CHROME SESSION DETECTED")
                print("Please complete login to HEWP in the Chrome window")
                print("After login, rerun this script to continue automation.")
                sys.exit(0)
            else:
                print("✅ Reconnected to existing Chrome session")
            print("="*50)
            print("STARTING AUTOMATION")
            print("="*50)
        except Exception as e:
            print(f"❌ Browser connection failed: {str(e)}")
            sys.exit(1)

    def ensure_window_visible(self):
        """Bring Chrome window to foreground if minimized"""
        print("[WINDOW] Ensuring Chrome window is visible...")
        try:
            chrome_windows = gw.getWindowsWithTitle("Chrome")
            if chrome_windows:
                chrome_window = chrome_windows[0]
                if chrome_window.isMinimized:
                    print("[WINDOW] Restoring minimized Chrome window.")
                    chrome_window.restore()
                chrome_window.activate()
                print("[WINDOW] Chrome window activated.")
                time.sleep(0.5)
            else:
                print("[WINDOW] No Chrome window found.")
        except Exception as e:
            print(f"Window management warning: {str(e)}")


    def ask_for_item_number(self):
        print("[INPUT] Asking for item number...")
        try:
            clipboard_text = pyperclip.paste()
        except Exception:
            clipboard_text = ""
        item_number = clipboard_text.strip()

        item_number_input = item_number
        print(f"[INPUT] Item number entered: {item_number_input}")
        if not item_number_input:
            return None
        return item_number_input

    def _get_valid_options(self, select_element):
        print("[DROPDOWN] Getting valid options from dropdown...")
        return [
            (opt.get_attribute("value"), opt.text.strip())
            for opt in select_element.find_elements(By.TAG_NAME, "option")
            if opt.get_attribute("value") and opt.text.strip() and opt.text.strip().lower() != "select one"
        ]

    def _prompt_dropdown(self, title, prompt, options):
        print(f"[PROMPT] Prompting user: {title} - {prompt}")
        root = tk.Tk()
        root.withdraw()
        win = tk.Toplevel(root)
        win.title(title)
        win.configure(bg="#f5f6fa")
        win.resizable(False, False)
        win.minsize(700, 260)
        win.attributes('-topmost', True)
        win.update_idletasks()
        width, height = 900, 400
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")
        style = ttk.Style(win)
        style.theme_use('clam')
        style.configure("TLabel", background="#f5f6fa", font=("Segoe UI", 12, "bold"), foreground="#222")
        style.configure("TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#0078D7")
        style.map("TButton",
              background=[('active', '#005fa3'), ('!active', '#0078D7')],
              foreground=[('active', '#fff'), ('!active', '#fff')])
        style.configure("TCombobox", font=("Segoe UI", 11), padding=6)
    
    
        # Tooltip for full text, always visible above the prompt label
        tooltip_var = tk.StringVar()
        tooltip = tk.Label(
            win,
            textvariable=tooltip_var,
            bg="#ffffe0",
            fg="#222",
            font=("Segoe UI", 10),
            relief="solid",
            borderwidth=1,
            wraplength=800,
            justify="left"
        )
        tooltip.pack(padx=40, pady=(10, 0), fill="x")  # Pack before the prompt label

        # Prompt label
        label = ttk.Label(win, text=prompt, anchor="center", style="TLabel")
        label.pack(padx=30, pady=(10, 10), fill="x")

        # Dropdown (wider for long text)
        var = tk.StringVar(win)
        var.set(options[0][1])
        combo = ttk.Combobox(
            win,
            textvariable=var,
            values=[o[1] for o in options],
            state="readonly",
            width=60,  # Wider for long text
            style="TCombobox"
        )
        combo.pack(padx=40, pady=10)
        combo.configure(justify="left")
        combo.current(0)

        def update_tooltip(event=None):
            idx = combo.current()
            if idx >= 0:
                tooltip_var.set(options[idx][1])
            else:
                tooltip_var.set("")

        combo.bind("<<ComboboxSelected>>", update_tooltip)
        update_tooltip()
    
        # --- Footer: OK & Cancel Buttons at the bottom ---
        result = {"value": None}
        def on_ok(event=None):

            idx = combo.current()
            result["value"] = options[idx][0]  # Return the value, not the text
            win.destroy()
            root.quit()
        def on_cancel(event=None):
            result["value"] = None
            win.destroy()
            root.quit()
            sys.exit(0)

        # Bind Enter and Esc keys
        win.bind('<Return>', on_ok)
        win.bind('<Escape>', on_cancel)
    
        # Buttons frame
        btn_frame = tk.Frame(win, bg="#f5f6fa")
        btn_frame.pack(side="bottom", pady=(10, 0))
        ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok, style="TButton")
        ok_btn.pack(side="left", padx=10)
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, style="TButton")
        cancel_btn.pack(side="left", padx=10)
        ok_btn.focus_set()
    
        # Watermark / Developer signature (still at the very bottom)
        watermark = tk.Label(
            win,
            text="Developed by MRGARGSIR",
            font=("Segoe UI", 10, "italic"),
            fg="#b0b0b0",
            bg="#f5f6fa",
            anchor="center"
        )
        watermark.pack(side="bottom", pady=(0, 10), fill="x")
    
        win.protocol("WM_DELETE_WINDOW", on_cancel)
        win.mainloop()
        root.destroy()
        print(f"[PROMPT] User selected: {result['value']}")
        return result["value"]

    def ensure_final_date(self):
        """If txtfinaldate field is present, ensure it has a value. Prompt user if not."""
        try:
            final_date_elem = self.driver.find_element(By.ID, "txtfinaldate")
            final_date_value = final_date_elem.get_attribute("value").strip()
            if not final_date_value:
                root = tk.Tk()
                root.withdraw()
                while True:
                    user_date = simpledialog.askstring(
                        "Bill Final Date Required",
                        "Enter Bill Final Date (DD/MM/YYYY):",
                        parent=root
                    )
                    if user_date is None:
                        messagebox.showerror("Input Required", "Bill Final Date is required to proceed.")
                        continue
                    user_date = user_date.strip()
                    import re
                    if re.match(r"^\d{2}/\d{2}/\d{4}$", user_date):
                        break
                    else:
                        messagebox.showerror("Invalid Format", "Please enter date in DD/MM/YYYY format.")
                root.destroy()
                self.driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                    final_date_elem, user_date
                )
                time.sleep(1)
        except Exception:
            pass  # Field not present, nothing to do


    
    def select_dropdowns_in_order(self, item_number=None):
        """Ensure all required dropdowns are selected in order, prompting user if needed.
        Only after the last dropdown, check ddlitemnumber for valid options."""

        dropdowns = [
            ("ddlcomp", "Main Component"),
            ("ddlsubhead", "Sub-Component"),
            
        ]
        idx = 0
        while idx < len(dropdowns):
            dropdown_id, label = dropdowns[idx]
            try:
                select_elem = self.driver.find_element(By.ID, dropdown_id)
            except Exception:
                idx += 1
                continue  # If dropdown not present, skip

            while True:
                valid_options = self._get_valid_options(select_elem)
                selected_value = select_elem.get_attribute("value")
                try:
                    selected_text = select_elem.find_element(By.CSS_SELECTOR, "option:checked").text.strip()
                except Exception:
                    selected_text = ""

               
                    time.sleep(1)
                    # After selecting, check if next dropdown has valid options
                    # WE CAN DELETE THIS FOR FAST RUN , ALREADY TESTED
                    if idx + 1 < len(dropdowns):
                        next_id, next_label = dropdowns[idx + 1]
                        try:
                            next_elem = self.driver.find_element(By.ID, next_id)
                            next_valid_options = self._get_valid_options(next_elem)
                            if not next_valid_options:
                                continue  # Prompt again for current
                        except Exception:
                            continue  # Prompt again for current
                    break  # Move to next dropdown

                need_prompt = not selected_value or selected_text.lower() == "select one" or not selected_text
                prompt_text = f"Select {label}:"
                if need_prompt:
                    if not valid_options:
                        break  # No valid options to select
                    chosen_value = self._prompt_dropdown(label, prompt_text, valid_options)
                    if chosen_value:
                        self.driver.execute_script(
                            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                            select_elem, chosen_value
                        )
                        time.sleep(1)
                    select_elem = self.driver.find_element(By.ID, dropdown_id)
                    continue  # Prompt again for current

                # For all except last dropdown, check next dropdown for valid options
                if idx + 1 < len(dropdowns):
                    next_id, next_label = dropdowns[idx + 1]
                    try:
                        next_elem = self.driver.find_element(By.ID, next_id)
                        next_valid_options = self._get_valid_options(next_elem)
                        if not next_valid_options:
                            prompt_text = f"Last selection was not valid!\nSelect {label}:"
                            chosen_value = self._prompt_dropdown(label, prompt_text, valid_options)
                            if chosen_value:
                                self.driver.execute_script(
                                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                    select_elem, chosen_value
                                )
                                time.sleep(1)
                            select_elem = self.driver.find_element(By.ID, dropdown_id)
                            continue  # Prompt again for current
                    except Exception:
                        prompt_text = f"Last selection was not valid!\nSelect {label}:"
                        chosen_value = self._prompt_dropdown(label, prompt_text, valid_options)
                        if chosen_value:
                            self.driver.execute_script(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                select_elem, chosen_value
                            )
                            time.sleep(1)
                        select_elem = self.driver.find_element(By.ID, dropdown_id)
                        continue  # Prompt again for current

                # Only after the last dropdown, check ddlitemnumber for valid options
                elif idx == len(dropdowns) - 1:
                    try:
                        rate_type_elem = self.driver.find_element(By.ID, "ddlitemnumber")
                        rate_type_valid_options = self._get_valid_options(rate_type_elem)
                        # --- NEW CONDITION: Only options containing item_number are valid ---
                        if item_number:
                            rate_type_valid_options = [
                                (val, txt) for val, txt in rate_type_valid_options
                                if item_number.lower() in txt.lower()
                            ]
                        if not rate_type_valid_options:
                            # Prompt for previous dropdown (the current one)
                            prompt_text = f"No matching item found!\nSelect {label} again:"
                            chosen_value = self._prompt_dropdown(label, prompt_text, valid_options)
                            if chosen_value:
                                self.driver.execute_script(
                                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                    select_elem, chosen_value
                                )
                                time.sleep(1)
                            select_elem = self.driver.find_element(By.ID, dropdown_id)
                            continue  # Prompt again for current
                    except Exception:
                        prompt_text = f"Last selection was not valid!\nSelect {label}:"
                        chosen_value = self._prompt_dropdown(label, prompt_text, valid_options)
                        if chosen_value:
                            self.driver.execute_script(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                select_elem, chosen_value
                            )
                            time.sleep(1)
                        select_elem = self.driver.find_element(By.ID, dropdown_id)
                        continue  # Prompt again for current

                break  # Only break if all checks above pass
            idx += 1

    def search_and_select_item(self, item_number=None):
        
        """Select ddlitemnumber using Selenium, handling all dropdown dependencies."""
        


        try:
            # Ensure all upper dropdowns are selected
            self.select_dropdowns_in_order(item_number)

            # Now handle ddlitemnumber dropdown
            while True:
                try:
                    # Always re-find the dropdown after any page reload
                    rate_type_select = self.driver.find_element(By.ID, "ddlitemnumber")
                except NoSuchElementException:
                    print("ddlitemnumber dropdown not found, skipping ddlitemnumber selection.")
                    return

                try:
                    options = self._get_valid_options(rate_type_select)
                except StaleElementReferenceException:
                    # If stale, re-find and retry
                    time.sleep(1)
                    continue
                if not options:
                    # No valid ddlitemnumber options, so re-run dropdown selection and try again
                    messagebox.showinfo(
                        "No Valid Rate Type",
                        "No valid Rate Type options available.\nPlease reselect previous dropdowns."
                    )
                    self.select_dropdowns_in_order()
                    continue  # Try again after reselection

                # Preferred order: clipboard/user ddlitemnumber, then defaults          

                selected = False
                if item_number:
                    # Partial match support (case-insensitive)
                    for value, text in options:
                        if item_number.lower() in text.lower():
                            self.driver.execute_script(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                rate_type_select, value
                            )
                            selected = True
                            break

                # If none of the preferred, select first valid option
                if not selected and options:
                    self.driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                        rate_type_select, options[0][0]
                    )

                # Highlight for feedback
                try:
                    self.driver.execute_script(
                        "arguments[0].style.outline='10px solid orange'; arguments[0].style.backgroundColor='#bf360c';",
                        rate_type_select
                    )
                    time.sleep(1)
                    self.driver.execute_script(
                        "arguments[0].style.outline=''; arguments[0].style.backgroundColor='';",
                        rate_type_select
                    )
                except StaleElementReferenceException:
                    # The element was replaced after selection, so just skip highlighting
                    pass
                break  # Exit loop if selection was successful
            print("[SELECT] Item selection complete.")
            
        except Exception as e:
            print(f"Error in select_rate_type_with_script: {e}")
            traceback.print_exc()
    
    
    def ensure_subhead_selected(self):
        """Ensure the '[Tender ID] & Agreement Name' (ddltender) dropdown is selected, prompting user if needed.
        If not present, navigate to the correct page using menu clicks."""
        
        # Ensure page is loaded and user is logged in
        print("[CHECK] Ensuring user is logged in and on the correct page...")
        user_elem = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.ID, "lblusername"))
        )
        # Ensure user is logged in by checking for the username label
        print("[CHECK] Checking if user is logged in...")
        try:
            user_elem = self.driver.find_element(By.ID, "lblusername")
            username = user_elem.text.strip()
            if not username:
                raise Exception("User not logged in or username not found.")
            else:
                print(f"[CHECK] User '{username}' is logged in.")
        except Exception:
            messagebox.showerror("Login Error", "You are not logged in. Please login first.")
            raise

        # Wait for ddltender dropdown to appear
        print("[CHECK] Ensuring '[Tender ID] & Agreement Name' dropdown is present...")
        try:
            select_elem = self.driver.find_element(By.ID, "ddltender")
        except Exception:
            # Wait up to 5 seconds for the dropdown to appear before deciding we're not on the right page
            

            try:
                select_elem = WebDriverWait(self.driver, 0.5).until(
                    EC.presence_of_element_located((By.ID, "ddltender"))
                )
            except TimeoutException:
                # Not on the right page, so navigate using menu
                print("[CHECK] '[Tender ID] & Agreement Name' dropdown not found, checking page navigation...")
                try:
                    print("[NAVIGATION] Ensuring Chrome window is visible...")
                    self.ensure_window_visible()  # Ensure Chrome window is visible

                    print("[NAVIGATION] Clicking sidebar menu: 'Submit Bill to JE'...")
                    # 1. Click the sidebar menu: <a href="#actmenucon" ...>Submit Bill to JE</a>
                    sidebar_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, '#actmenucon') and contains(text(), 'Submit Bill to JE')]")
                    sidebar_link = next((l for l in sidebar_links if l.is_displayed()), None)
                    if sidebar_link:
                        self.driver.execute_script("arguments[0].scrollIntoView();", sidebar_link)
                        try:
                            sidebar_link.click()
                        except (ElementClickInterceptedException, ElementNotInteractableException):
                            self.driver.execute_script("arguments[0].click();", sidebar_link)
                        print("[NAVIGATION] Sidebar menu clicked.")
                        time.sleep(1)
                    else:
                        print("[ERROR] Sidebar 'Submit Bill to JE' menu not found.")
                        raise Exception("Sidebar 'Submit Bill to JE' menu not found.")

                    print("[NAVIGATION] Clicking submenu: 'Submit Bill to JE'...")
                    # 2. Click submenu: <a ... href="/E-Billing/Est_Add_Items_emb.aspx">Submit Bill to JE</a>
                    submenu_links = self.driver.find_elements(
                        By.XPATH,
                        "//ul[contains(@id, 'actmenucon')]/li/a[contains(@href, '/E-Billing/Est_Add_Items_emb.aspx') and contains(text(), 'Submit Bill to JE')]"
                    )
                    submenu_link = next((l for l in submenu_links if l.is_displayed()), None)
                    if submenu_link:
                        self.driver.execute_script("arguments[0].scrollIntoView();", submenu_link)
                        try:
                            submenu_link.click()
                        except (ElementClickInterceptedException, ElementNotInteractableException):
                            self.driver.execute_script("arguments[0].click();", submenu_link)
                        print("[NAVIGATION] Submenu clicked. Waiting for page to load...")
                        time.sleep(2)
                    else:
                        print("[ERROR] Submenu 'Submit Bill to JE' link not found.")
                        raise Exception("Submenu 'Submit Bill to JE' link not found.")
                except Exception as nav_err:
                    print(f"[ERROR] Could not navigate to Add Items in Bill page: {nav_err}")
                    messagebox.showerror("Navigation Error", f"Could not navigate to Add Items in Bill page:\n{nav_err}")
                    raise
                # Wait for ddltender to appear
                # Try again to find the dropdown after navigation
                try:
                    print("[NAVIGATION] Waiting for '[Tender ID] & Agreement Name' dropdown to appear after navigation...")
                    select_elem = self.wait.until(EC.presence_of_element_located((By.ID, "ddltender")))
                    print("[NAVIGATION] '[Tender ID] & Agreement Name' dropdown found.")
                except Exception:
                    print("[ERROR] Could not find '[Tender ID] & Agreement Name' dropdown after navigation.")
                    messagebox.showerror("Navigation Error", f"Could not navigate to Add Items in Bill. Please navigate manually.")
                    raise
        
        # Now prompt user if not selected
        valid_options = self._get_valid_options(select_elem)
        selected_value = select_elem.get_attribute("value")
        try:
            selected_text = select_elem.find_element(By.CSS_SELECTOR, "option:checked").text.strip()
        except Exception:
            selected_text = ""

        if not selected_value or selected_text.lower() == "select one" or not selected_text:
            if not valid_options:
                return  # No valid options to select
            chosen_value = self._prompt_dropdown("[Tender ID] & Agreement Name", "Select [Tender ID] & Agreement Name:", valid_options)
            if chosen_value:
                self.driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                    select_elem, chosen_value
                )
                time.sleep(1)  # Wait for page reload

       

    def process_item(self):
        """Main workflow with window management"""
        try:
            item_number = self.ask_for_item_number()
            if not item_number:
                return
            self.ensure_subhead_selected()  # Ensure ddltender is selected before item selection
            self.ensure_final_date()  # Ensure final date is set if applicable
            self.search_and_select_item(item_number)
            
            print("[PROCESS] Item selection complete. ..")
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            return False
        finally:
            pass

    def close(self):
        """Clean up resources"""
        if hasattr(self, 'driver'):
            # self.driver.quit()
            print("skipped Closing browser session...")
            pass
        if self.notification_root:
            try:
                self.notification_root.destroy()
            except:
                pass

def run_automation():
    """Run the automation process"""
    uploader = HEWPSetItem()
    try:
        uploader.process_item()
    finally:
        uploader.close()

if __name__ == "__main__":
    run_automation()
    sys.exit(0)  # Ensures the script exits cleanly