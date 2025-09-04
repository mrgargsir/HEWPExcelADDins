import os
import sys
import ctypes
import time
import subprocess
import importlib.util

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pyperclip
import tkinter as tk
from threading import Timer
from tkinter import messagebox
import pygetwindow as gw
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException, TimeoutException

class HEWPUploader:
    def __init__(self):
        self.notification_root = None
        self._chrome_was_launched = False

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
            "selenium", "pyautogui", "pyperclip", "pygetwindow", "tkinter"
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

    def auto_close_info(self, title, message, timeout=500):
        root = tk.Tk()
        root.withdraw()
        win = tk.Toplevel(root)
        win.title(title)
        win.configure(bg="#f5f6fa")
        win.resizable(False, False)
        win.attributes('-topmost', True)
        label = tk.Label(win, text=message, font=("Segoe UI", 13, "bold"), bg="#f5f6fa", fg="#222")
        label.pack(padx=30, pady=30)
        # Center window
        win.update_idletasks()
        width, height = win.winfo_width(), win.winfo_height()
        # x = (win.winfo_screenwidth() // 2) - (width // 2)
        # y = (win.winfo_screenheight() // 2) - (height // 2)
        x = (win.winfo_screenwidth()) - (width) - 10
        y = ((win.winfo_screenheight()//7) *6) - (height ) -10
        win.geometry(f"+{x}+{y}")
        # Auto-close after timeout ms
        win.after(timeout, win.destroy)
        win.mainloop()
        root.destroy()

    def ask_for_item_number(self):
        """Prompt user to enter item number (shows clipboard content as default).
        Clipboard format: [item_number:rate_type
        Returns: (item_number, rate_type)
        """
        try:
            clipboard_text = pyperclip.paste()
        except Exception:
            clipboard_text = ""

        # Parse clipboard for [item_number:rate_type
        item_number = ""
        rate_type = ""
        if clipboard_text.startswith("[") and ":" in clipboard_text:
            try:
                # Remove leading '[' and split by ':'
                parts = clipboard_text[1:].split(":", 1)
                item_number = parts[0].strip()
                rate_type = parts[1].strip() if len(parts) > 1 else ""
            except Exception:
                item_number = clipboard_text.strip()
                rate_type = ""
        else:
            item_number = clipboard_text.strip()

        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.withdraw()
        win = tk.Toplevel(root)
        win.title("Item Number Input")
        win.configure(bg="#f5f6fa")
        win.resizable(False, False)
        win.minsize(440, 200)

        # Center the window on the screen
        win.update_idletasks()
        width = 440
        height = 300
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

        # Premium style
        style = ttk.Style(win)
        style.theme_use('clam')
        style.configure("TLabel", background="#f5f6fa", font=("Segoe UI", 12, "bold"), foreground="#222")
        style.configure("TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#0078D7")
        style.map("TButton",
                  background=[('active', '#005fa3'), ('!active', '#0078D7')],
                  foreground=[('active', '#fff'), ('!active', '#fff')])
        style.configure("TEntry", font=("Segoe UI", 12), padding=6)

        # Prompt label
        label = ttk.Label(win, text="Enter the item number (can be partial):", anchor="center", style="TLabel")
        label.pack(padx=30, pady=(30, 10), fill="x")

        # Entry box
        var = tk.StringVar(win)
        var.set(item_number)
        entry = ttk.Entry(win, textvariable=var, width=36, style="TEntry", justify="center")
        entry.pack(padx=40, pady=10)
        entry.focus_set()
        

        # OK & Cancel Buttons
        result = {"value": None}
       
        def on_ok(event=None):
            result["value"] = var.get().strip()
            win.destroy()
            root.quit()
        def on_cancel(event=None):
            result["value"] = None
            win.destroy()
            root.quit()
            sys.exit(0)
       

        entry.bind('<Return>', on_ok)

        # Bind Enter and Esc keys
        win.bind('<Return>', on_ok)
        win.bind('<Escape>', on_cancel)

        # Footer frame for buttons
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

        item_number_input = result["value"]
        if not item_number_input:
            return None, None
        return item_number_input, rate_type

    def search_and_select_item(self, item_number):
        """Search and select item that contains the number, supporting both dropdown and textbox+button.
        Uses id, name, and value for robust element matching."""
        try:
            # Try dropdown first (check both id and name, and non-blank options)
            dropdown = None
            try:
                dropdown = self.driver.find_element(By.CSS_SELECTOR, "select#ddlitemnumber[name*='ddlitemnumber']")
                if not (dropdown.get_attribute("id") == "ddlitemnumber" and "ddlitemnumber" in dropdown.get_attribute("name")):
                    dropdown = None
            except Exception:
                dropdown = None

            dropdown_valid = False
            if dropdown:
                # Check for at least one non-blank, non-"Select One" option
                options = [
                    opt for opt in dropdown.find_elements(By.TAG_NAME, "option")
                    if opt.get_attribute("value") and opt.text.strip() and opt.text.strip().lower() != "select one"
                ]
                if options:
                    dropdown_valid = True

            if dropdown_valid:
                # Use JS for search box if present
                self.driver.execute_script("""
                    if (!document.getElementById('dynamicSearchContainer')) {
                        const container = document.createElement('div');
                        container.id = 'dynamicSearchContainer';
                        container.style.position = 'relative';
                        container.style.margin = '10px 0';
                        const searchInput = document.createElement('input');
                        searchInput.id = 'dynamicSearchInput';
                        searchInput.type = 'text';
                        searchInput.placeholder = '🔍 Enter item number...';
                        searchInput.style.width = '100%';
                        searchInput.style.padding = '8px';
                        document.querySelector('#ddlitemnumber').parentNode.prepend(container);
                        container.appendChild(searchInput);
                    }
                    document.getElementById('dynamicSearchInput').value = arguments[0];
                """, item_number)
                # Select the item in dropdown
                for option in options:
                    if item_number in option.text:
                        option.click()
                        break
                else:
                    raise ValueError(f"Item '{item_number}' not found in dropdown")
                time.sleep(1)
                return  # Success, exit function

            # Fallback to textbox+button (always available)
            txthsrno = self.driver.find_element(By.CSS_SELECTOR, "input#txthsrno[name*='txthsrno']")
            # Find all submit buttons and match by id, name, and value
            buttons = self.driver.find_elements(By.CSS_SELECTOR, "input[type='submit']")
            button = None
            for btn in buttons:
                if (
                    btn.get_attribute("id") == "Button5"
                    and "Button5" in btn.get_attribute("name")
                    and btn.get_attribute("value").strip().lower() == "search item"
                ):
                    button = btn
                    break
            if txthsrno.get_attribute("id") == "txthsrno" and "txthsrno" in txthsrno.get_attribute("name") and button:
                txthsrno.clear()
                txthsrno.send_keys(item_number)
                button.click()
                time.sleep(1)
                return  # Success, exit function

            # If neither worked, show error
            messagebox.showinfo(
                "Not Ready",
                "Neither item dropdown nor item textbox+button found on page.\n\nPlease reach the destination page on website."
            )
            raise RuntimeError("No item selection UI found on page.")

        except RuntimeError as e:
            # End automation if selection UI is not found
            messagebox.showerror("Automation Stopped", str(e))
            self.close()
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("Selection Error", f"Could not select item: {str(e)}")
            raise

    def _get_valid_options(self, select_element):
        """Return list of (value, text) tuples for valid options (not blank or 'Select One')."""
        return [
            (opt.get_attribute("value"), opt.text.strip())
            for opt in select_element.find_elements(By.TAG_NAME, "option")
            if opt.get_attribute("value") and opt.text.strip() and opt.text.strip().lower() != "select one"
        ]

    def _prompt_dropdown(self, title, prompt, options):
        """Prompt user to select from options using a premium-looking, centered dropdown dialog with watermark and Cancel button."""
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.withdraw()
        win = tk.Toplevel(root)
        win.title(title)
        win.configure(bg="#f5f6fa")
        win.resizable(False, False)
        win.minsize(700, 260)  # Wider window for long text

        # Center the window on the screen
        win.update_idletasks()
        width = 900
        height = 400
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

        # Premium style
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

        # Footer frame for buttons
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
        return result["value"]

    def select_dropdowns_in_order(self):
        """Ensure all required dropdowns are selected in order, prompting user if needed.
        Only after the last dropdown, check Rate_Type for valid options."""
        dropdowns = [
            ("ddlchapter", "HSR Chapter Name"),
            ("ddlPremiumDate", "Premium Date"),
            ("ddlclass", "Class Name"),
            ("ddlsection", "HSR Section Name"),
            ("ddlsubsection", "HSR Sub Section Name"),
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

                # For Premium Date, always select the latest (max) date value (format: DD/MM/YYYY)
                if dropdown_id == "ddlPremiumDate" and valid_options:
                    from datetime import datetime
                    def parse_date(text):
                        try:
                            return datetime.strptime(text, "%d/%m/%Y")
                        except Exception:
                            return datetime.min
                    latest_option = max(valid_options, key=lambda opt: parse_date(opt[1]))
                    latest_value = latest_option[0]
                    self.driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                        select_elem, latest_value
                    )
                    time.sleep(1)
                    # After selecting, check if next dropdown has valid options
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

                # Only after the last dropdown, check Rate_Type for valid options
                elif idx == len(dropdowns) - 1:
                    try:
                        rate_type_elem = self.driver.find_element(By.ID, "ddlRate_Type")
                        rate_type_valid_options = self._get_valid_options(rate_type_elem)
                        if not rate_type_valid_options:
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

                break  # Only break if all checks above pass
            idx += 1

    def select_rate_type_with_script(self, rate_type=None):
        """Select Rate Type using Selenium, handling all dropdown dependencies."""
        from selenium.common.exceptions import NoSuchElementException
        import time

        try:
            # Ensure all upper dropdowns are selected
            self.select_dropdowns_in_order()

            # Now handle Rate Type dropdown
            while True:
                try:
                    # Always re-find the dropdown after any page reload
                    rate_type_select = self.driver.find_element(By.ID, "ddlRate_Type")
                except NoSuchElementException:
                    print("Rate_Type dropdown not found, skipping Rate Type selection.")
                    return

                options = self._get_valid_options(rate_type_select)
                if not options:
                    # No valid Rate Type options, so re-run dropdown selection and try again
                    messagebox.showinfo(
                        "No Valid Rate Type",
                        "No valid Rate Type options available.\nPlease reselect previous dropdowns."
                    )
                    self.select_dropdowns_in_order()
                    continue  # Try again after reselection

                # Preferred order: clipboard/user rate_type, then defaults
                preferred = [rate_type] if rate_type else []
                preferred += ["Through Rate", "Rate", "Labour Rate"]

                selected = False
                for pref in preferred:
                    for value, text in options:
                        if text == pref:
                            self.driver.execute_script(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                                rate_type_select, value
                            )
                            selected = True
                            break
                    if selected:
                        break

                # If none of the preferred, select first valid option
                if not selected and options:
                    self.driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                        rate_type_select, options[0][0]
                    )

                # Highlight for feedback
                self.driver.execute_script(
                    "arguments[0].style.outline='3px solid orange'; arguments[0].style.backgroundColor='#bf360c';",
                    rate_type_select
                )
                time.sleep(1)
                self.driver.execute_script(
                    "arguments[0].style.outline=''; arguments[0].style.backgroundColor='';",
                    rate_type_select
                )
                break  # Exit loop if selection was successful
            print("[SELECT] Item and Rate Type selection complete.")
            self.auto_close_info("Item Selection", "Item selection complete.", timeout=1000)
        except Exception as e:
            print(f"Error in select_rate_type_with_script: {e}")

    def ensure_subhead_selected(self):
        """Ensure the 'Name of Template' (ddlsubhead) dropdown is selected, prompting user if needed.
        If not present, navigate to the correct page using menu clicks."""

        print("[CHECK] Ensuring user is logged in and on the correct page...")
        current_url = self.driver.current_url
        print("[DEBUG] Current URL:", current_url)

        # If not on the correct domain, try to find the right tab
        if "works.haryana.gov.in" not in current_url:
            print("[SWITCH] Looking for HEWP tab...")
            for handle in self.driver.window_handles:
                self.driver.switch_to.window(handle)
                url = self.driver.current_url
                print("[SWITCH] Tab URL:", url)
                if "works.haryana.gov.in" in url:
                    print("[SWITCH] Found HEWP tab.")
                    break
            else:
                print("[NAVIGATION] Navigating to HEWP portal home page...")
                self.driver.get("https://works.haryana.gov.in/contractor/contractorHome.aspx")
                time.sleep(2)
                current_url = self.driver.current_url
                print("[DEBUG] New URL after navigation:", current_url)
                if "works.haryana.gov.in" not in current_url:
                    messagebox.showerror(
                        "Navigation Error",
                        "Could not navigate to the HEWP portal.\nPlease open https://works.haryana.gov.in/contractor/contractorHome.aspx in Chrome and log in."
                    )
                    raise Exception("Not on HEWP portal.")

        # 2. Check for login status
        try:
            user_elem = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.ID, "lblusername"))
            )
            username = user_elem.text.strip()
            if not username:
                raise Exception("User not logged in or username not found.")
            print(f"[CHECK] User '{username}' is logged in.")
        except Exception:
            # Try to detect login form as fallback
            try:
                login_input = self.driver.find_element(By.XPATH, "//input[@type='password' or contains(@id, 'login')]")
                messagebox.showerror("Login Error", "You are not logged in. Please log in to HEWP portal in Chrome and try again.")
            except Exception:
                messagebox.showerror("Login Error", "Could not detect login status. Please ensure you are logged in to HEWP portal.")
            raise

        # 3. Now continue as before...

        try:
            select_elem = self.driver.find_element(By.ID, "ddlsubhead")
        except Exception:
            # Wait up to 5 seconds for the dropdown to appear before deciding we're not on the right page
            
            try:
                select_elem = WebDriverWait(self.driver, 0.5).until(
                    EC.presence_of_element_located((By.ID, "ddlsubhead"))
                )
            except TimeoutException:
                # Not on the right page, so navigate using menu
                print("[NAVIGATION] Not on the Add/Edit Items page, navigating...")

                try:
                    self.ensure_window_visible()  # Ensure Chrome window is visible
                    # 1. Click the visible "e-Estimate" menu (find the one that's displayed)
                    e_estimate_links = self.driver.find_elements(By.XPATH, "//a[@id='HyperLink7' and contains(text(),'e-Estimate')]")
                    e_estimate = None
                    for link in e_estimate_links:
                        if link.is_displayed():
                            e_estimate = link
                            break
                    if not e_estimate:
                        raise Exception("Could not find visible 'e-Estimate' menu link.")
                    self.driver.execute_script("arguments[0].scrollIntoView();", e_estimate)
                    try:
                        e_estimate.click()
                    except (ElementClickInterceptedException, ElementNotInteractableException):
                        self.driver.execute_script("arguments[0].click();", e_estimate)
                    time.sleep(1)

                    # 2. Wait for "Add Estimate Template" to be visible and click it
                    add_template = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'Est_Template_Add_Name.aspx')]"))
                    )
                    self.driver.execute_script("arguments[0].scrollIntoView();", add_template)
                    try:
                        add_template.click()
                    except (ElementClickInterceptedException, ElementNotInteractableException):
                        self.driver.execute_script("arguments[0].click();", add_template)
                    time.sleep(1)

                    # 3. Wait for "Add / Edit Items" to be visible and click it
                    add_edit_items = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'Est_Template_Add_Items.aspx')]"))
                    )
                    self.driver.execute_script("arguments[0].scrollIntoView();", add_edit_items)
                    try:
                        add_edit_items.click()
                    except (ElementClickInterceptedException, ElementNotInteractableException):
                        self.driver.execute_script("arguments[0].click();", add_edit_items)
                    time.sleep(2)
                except Exception as nav_err:
                    messagebox.showerror("Navigation Error", f"Could not navigate to Add/Edit Items page:\n{nav_err}")
                    raise

            # Try again to find the dropdown after navigation
            try:
                select_elem = self.driver.find_element(By.ID, "ddlsubhead")
            except Exception:
                messagebox.showerror("Page Error", "Could not reach the Add/Edit Items page. Please navigate manually.")
                raise

        valid_options = self._get_valid_options(select_elem)
        selected_value = select_elem.get_attribute("value")
        try:
            selected_text = select_elem.find_element(By.CSS_SELECTOR, "option:checked").text.strip()
        except Exception:
            selected_text = ""

        # If not selected or selected is blank/"Select One", prompt user
        if not selected_value or selected_text.lower() == "select one" or not selected_text:
            if not valid_options:
                return  # No valid options to select
            chosen_value = self._prompt_dropdown("Name of Template", "Select Name of Template:", valid_options)
            if chosen_value:
                self.driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                    select_elem, chosen_value
                )
                time.sleep(1)  # Wait for page reload

    def process_item(self):
        """Main workflow with window management"""
        try:
            item_number, rate_type = self.ask_for_item_number()
            if not item_number:
                return
            self.ensure_subhead_selected()  # Ensure ddlsubhead is selected before item selection
            self.ensure_window_visible()  # Ensure Chrome window is visible
            self.search_and_select_item(item_number)
            self.select_rate_type_with_script(rate_type)  # Pass rate_type here

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
    uploader = HEWPUploader()
    try:
        uploader.process_item()
    finally:
        uploader.close()

if __name__ == "__main__":
    run_automation()
    sys.exit(0)  # Ensures the script exits cleanly