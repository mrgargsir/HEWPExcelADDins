import os
import sys
import ctypes
import time
import subprocess
import importlib.util
import pandas as pd
import math
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pyperclip
import tkinter as tk
from tkinter import Toplevel, Label, messagebox
import pygetwindow as gw

class HEWPwritter:
    def __init__(self):
        print("[INIT] Initializing HEWPwritter...")
        
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

    def ensure_excel_window_visible(self):
        """Ensure Excel window is open and visible, else show error and exit."""
        print("[EXCEL] Checking for Excel window...")
        try:
            excel_windows = gw.getWindowsWithTitle("Excel")
            if not excel_windows:
                messagebox.showerror("Excel Not Found", "Please open your Excel file and select the data before running this script.")
                print("[EXCEL] Excel window not found. Please open Excel and select your data.")
                sys.exit(1)
            excel_window = excel_windows[0]
            if excel_window.isMinimized:
                print("[EXCEL] Excel window is minimized. Restoring...")
                excel_window.restore()
            excel_window.activate()
            print("[EXCEL] Excel window activated and brought to foreground.")
            time.sleep(0.5)
        except Exception as e:
            print(f"[EXCEL] Error ensuring Excel window is visible: {e}")
            messagebox.showerror("Excel Error", f"Error ensuring Excel window is visible:\n{e}")
            sys.exit(1)

    def fill_portal_row(self, row_data, row_index):
        driver = self.driver
        wait = self.wait
        suffix = f"_ctl{row_index}"

        print(f"[DEBUG] Filling portal row: index={row_index}, data={row_data}")

        try:
            print(f"[DEBUG] Locating unit element for row {row_index}...")
            unit_elem = driver.find_element(By.ID, 'unit')
            unit = unit_elem.text.strip().lower()
            print(f"[DEBUG] Unit for row {row_index}: {unit}")

            if 'Description' in row_data:
                desc_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtdescription"
                print(f"[DEBUG] Filling Description for row {row_index} using name {desc_name}")
                desc_elem = wait.until(EC.presence_of_element_located((By.NAME, desc_name)))
                desc_elem.clear()
                desc_elem.send_keys(str(row_data.get('Description', '')))
                print(f"[DEBUG] Description filled: {row_data.get('Description', '')}")

                # --- Set Plus/Minus dropdown based on product sign ---
                try:
                    num = float(row_data.get('Number', 1) or 1)
                    length = float(row_data.get('Length', 1) or 1)
                    breadth = float(row_data.get('Breadth', 1) or 1)
                    depth = float(row_data.get('Depth', 1) or 1)
                    product = num * length * breadth * depth
                    ddl_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:ddlplusminus"
                    ddl_elem = driver.find_element(By.NAME, ddl_name)
                    if product < 0:
                        print(f"[DEBUG] Product is negative ({product}), setting dropdown to Minus.")
                        ddl_elem.send_keys('-')
                    else:
                        print(f"[DEBUG] Product is positive ({product}), keeping dropdown as Plus.")
                        ddl_elem.send_keys('+')
                except Exception as e:
                    print(f"[DEBUG] Could not set Plus/Minus dropdown: {e}")

            if 'Number' in row_data:
                num_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtNumber"
                print(f"[DEBUG] Filling Number for row {row_index} using name {num_name}")
                num_elem = driver.find_element(By.NAME, num_name)
                num_elem.clear()
                num_elem.send_keys(str(row_data.get('Number', '')))
                print(f"[DEBUG] Number filled: {row_data.get('Number', '')}")

            if unit == 'cum':
                if 'Length' in row_data:
                    len_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtLength"
                    print(f"[DEBUG] Filling Length for row {row_index} using name {len_name}")
                    len_elem = driver.find_element(By.NAME, len_name)
                    len_elem.clear()
                    len_elem.send_keys(str(row_data.get('Length', '')))
                    print(f"[DEBUG] Length filled: {row_data.get('Length', '')}")
                if 'Breadth' in row_data:
                    br_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtBreadth"
                    print(f"[DEBUG] Filling Breadth for row {row_index} using name {br_name}")
                    br_elem = driver.find_element(By.NAME, br_name)
                    br_elem.clear()
                    br_elem.send_keys(str(row_data.get('Breadth', '')))
                    print(f"[DEBUG] Breadth filled: {row_data.get('Breadth', '')}")
                if 'Depth' in row_data:
                    dep_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtDepth"
                    print(f"[DEBUG] Filling Depth for row {row_index} using name {dep_name}")
                    dep_elem = driver.find_element(By.NAME, dep_name)
                    dep_elem.clear()
                    dep_elem.send_keys(str(row_data.get('Depth', '')))
                    print(f"[DEBUG] Depth filled: {row_data.get('Depth', '')}")
            elif unit == 'sqm':
                if 'Length' in row_data:
                    len_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtLength"
                    print(f"[DEBUG] Filling Length for row {row_index} using name {len_name}")
                    len_elem = driver.find_element(By.NAME, len_name)
                    len_elem.clear()
                    len_elem.send_keys(str(row_data.get('Length', '')))
                    print(f"[DEBUG] Length filled: {row_data.get('Length', '')}")
                if 'Breadth' in row_data:
                    br_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtBreadth"
                    print(f"[DEBUG] Filling Breadth for row {row_index} using name {br_name}")
                    br_elem = driver.find_element(By.NAME, br_name)
                    br_elem.clear()
                    br_elem.send_keys(str(row_data.get('Breadth', '')))
                    print(f"[DEBUG] Breadth filled: {row_data.get('Breadth', '')}")
            elif unit == 'rm':
                if 'Length' in row_data:
                    len_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtLength"
                    print(f"[DEBUG] Filling Length for row {row_index} using name {len_name}")
                    len_elem = driver.find_element(By.NAME, len_name)
                    len_elem.clear()
                    len_elem.send_keys(str(row_data.get('Length', '')))
                    print(f"[DEBUG] Length filled: {row_data.get('Length', '')}")
            else:
                if 'Quantity' in row_data:
                    qty_name = f"_ctl0:maincontentcm:GV_ADD_to_List:{suffix}:txtqty"
                    print(f"[DEBUG] Filling Quantity for row {row_index} using name {qty_name}")
                    qty_elem = driver.find_element(By.NAME, qty_name)
                    qty_elem.clear()
                    qty_elem.send_keys(str(row_data['Quantity']))
                    print(f"[DEBUG] Quantity filled: {row_data['Quantity']}")
        except Exception as e:
            print(f"[ERROR] Failed to fill portal row {row_index}: {e!r}")
            import traceback
            traceback.print_exc()
            raise

        print(f"[PORTAL] Data filled for row {row_index}")

    def clear_portal_fields(self):
        """Clear all portal input fields after processing."""
        print("[CLEAR] Clearing all portal fields...")
        try:
            # Adjust selectors as per your portal's actual field names/IDs
            for suffix in range(2, 12):  # Assuming up to 10 rows (2-11)
                row_suffix = f"_ctl{suffix}"
                try:
                    desc = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtdescription")
                    desc.clear()
                except Exception: pass
                try:
                    ddl = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:ddlplusminus")
                    ddl.send_keys('+')  # Reset to default Plus
                except Exception: pass
                try:
                    num = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtNumber")
                    num.clear()
                except Exception: pass
                try:
                    length = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtLength")
                    length.clear()
                except Exception: pass
                try:
                    breadth = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtBreadth")
                    breadth.clear()
                except Exception: pass
                try:
                    depth = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtDepth")
                    depth.clear()
                except Exception: pass
                try:
                    qty = self.driver.find_element(By.NAME, f"_ctl0:maincontentcm:GV_ADD_to_List:{row_suffix}:txtqty")
                    qty.clear()
                except Exception: pass

            print("[CLEAR] Portal fields cleared.")
        except Exception as e:
            print(f"[CLEAR] Error clearing portal fields: {e}")

    def load_selected_excel_data_no_headers(self):
        self.ensure_excel_window_visible()
        print("[EXCEL] Simulating Ctrl+C in Excel...")
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        print("[EXCEL] Reading data from clipboard (no headers)...")
        try:
            df = pd.read_clipboard(sep='\t', header=None)
            expected_cols = ['Description', 'Number', 'Length', 'Breadth', 'Depth']
            df.columns = expected_cols[:df.shape[1]]
            print(f"[EXCEL] DataFrame loaded with columns: {df.columns.tolist()} and {len(df)} rows.")

            processed_rows = []
            for idx, row in df.iterrows():
                print(f"[EXCEL] Processing row {idx+1}: {row.values}")
                desc = str(row['Description']).strip() if pd.notna(row['Description']) else ""
                if not desc:
                    desc = "."
                num = row['Number']
                if pd.isna(num):
                    print(f"[EXCEL] Skipping row {idx+1}: Number is blank.")
                    continue
                try:
                    num_val = float(num)
                except Exception:
                    print(f"[EXCEL] Skipping row {idx+1}: Number is not a valid float.")
                    continue
                if num_val == 0:
                    print(f"[EXCEL] Skipping row {idx+1}: Number is zero.")
                    continue

                # Shift non-zero, non-blank values left for Length, Breadth, Depth
                values = []
                for col in ['Length', 'Breadth', 'Depth']:
                    val = row[col]
                    try:
                        val_num = float(val)
                        # Check for nan and zero
                        if not math.isnan(val_num) and val_num != 0:
                            values.append(val_num)
                    except Exception:
                        pass
                while len(values) < 3:
                    values.append(0)
                values = values[:3]
                print(f"[EXCEL] Row {idx+1} processed as: Description={desc}, Number={num_val}, Length={values[0]}, Breadth={values[1]}, Depth={values[2]}")

                processed_rows.append({
                    'Description': desc,
                    'Number': int(num_val) if num_val.is_integer() else num_val,
                    'Length': values[0],
                    'Breadth': values[1],
                    'Depth': values[2]
                })

            self.excel_rows = processed_rows
            print(f"[EXCEL] Loaded {len(self.excel_rows)} processed rows from clipboard selection (no headers).")
        except Exception as e:
            print(f"[EXCEL] Failed to read clipboard data: {e}")
            self.excel_rows = []

    def submit_data(self): 
        """Click 'Add Items to List' button after filling row data"""
        driver = self.driver
        wait = self.wait
        print("[SUBMIT] Attempting to submit data to portal...")
        try:
            add_btn = driver.find_element(By.ID, 'btn_add_Description')
            print("[SUBMIT] Found 'Add Items to List' button. Clicking...")
            add_btn.click()
            # wait.until(EC.invisibility_of_element_located((By.ID, 'loadingSpinner')))
            time.sleep(0.5)  # Wait for any animations to finish
            print("[SUBMIT] Row added to portal.")
            print("[SUBMIT] Row submitted to portal.")
        except Exception as e:
            print(f"[SUBMIT] Failed to submit row: {e}")

    def handle_confirmation_and_scrolling(self):
        """EXACT implementation as you specified"""
        print("[CONFIRM] Handling confirmation and scrolling...")
        self.ensure_window_visible()
        try:
            # Handle SweetAlert error and summary if present

            # 1. First confirm any Chrome alert with Enter
            try:
                print("[CONFIRM] Checking for browser alert...")
                alert = self.wait.until(lambda d: d.switch_to.alert)
                alert.accept()
                print("[CONFIRM] Browser alert accepted.")
                time.sleep(0.5)
            except:
                print("[CONFIRM] No browser alert found.")
                pass
            # 2. Execute your exact JavaScript sequence
            print("[CONFIRM] Executing JavaScript for confirmation and scrolling.")
            self.driver.execute_script("""
                (function(){
                    const wait = ms => new Promise(r => setTimeout(r, ms));
                    
                    (async function(){
                        // Handle OK button
                        const okBtn = document.querySelector('button.confirm');
                        if (okBtn && okBtn.textContent.trim().toLowerCase() === 'ok') {
                            okBtn.click();
                            await wait(500);
                        }
                        
                        // Close modal
                        document.querySelector('button[data-bs-dismiss="modal"]')?.click();
                        await wait(500);
                        
                        // Handle rblitemshsr_1
                        const r = document.getElementById('rblitemshsr_1');
                        if (r) {
                            r.click();
                            r.scrollIntoView({behavior: 'instant', block: 'start'});
                            await wait(200);
                        }
                        
                        // Scroll to Description Details
                        const hs = document.querySelectorAll('.cust-card-heading h4');
                        for (const h of hs) {
                            const t = h.textContent.trim();
                            if (t === 'Description Details' || t.includes('Description Details')) {
                                h.scrollIntoView({behavior: 'instant', block: 'start'});
                                break;
                            }
                        }
                    })();
                })();
            """)
            print("[CONFIRM] Confirmation and scrolling completed.")
            time.sleep(0.5)
        except Exception as e:
            print(f"Warning during confirmation handling: {str(e)}")

    def handle_quantity_error_and_summary(self):
        """Handle SweetAlert error for excess quantity and show summary message."""
        print("[SUMMARY] Checking for quantity error and summary...")
        try:
            print("[SUMMARY] Waiting for SweetAlert error modal...")
            # Wait for the SweetAlert error modal to appear
            short_wait = WebDriverWait(self.driver, 5)
            error_modal = short_wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, '.sweet-alert.showSweetAlert.visible'))
            )
            print("[SUMMARY] Error modal found.")
            heading = error_modal.find_element(By.TAG_NAME, 'h2')
            print(f"[SUMMARY] Modal heading: '{heading.text.strip()}'")
            if heading.text.strip().lower() == "error":
                print("[SUMMARY] Quantity error detected. Handling error modal...")
                # Click the OK button in the error modal
                ok_btn = error_modal.find_element(By.CSS_SELECTOR, 'button.confirm')
                print("[SUMMARY] Clicking OK button in error modal.")
                ok_btn.click()
                time.sleep(0.5)
                # Click the "Clear Data" button
                print("[SUMMARY] Waiting for Clear Data button...")
                copy_btn = self.driver.find_elements(By.ID, "btnclear")
                if copy_btn and copy_btn[0].is_displayed():
                    copy_btn[0].click()
                    # Wait for the close button to be visible and clickable
                    close_btn = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.theme-btn--red[data-bs-dismiss="modal"]'))
                    )
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", close_btn)
                    time.sleep(0.1)  # Let any animation finish
                    close_btn.click()

                time.sleep(0.5)  # Wait for page to reload

                close_btn = self.driver.find_elements(By.CSS_SELECTOR, 'button.btn.theme-btn--red[data-bs-dismiss="modal"]')
                if close_btn and close_btn[0].is_displayed():
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", close_btn[0])
                    time.sleep(0.1)  # Let any animation finish
                    close_btn[0].click()

                # Append the summary message at the top
                print("[SUMMARY] Appending summary message to page.")
                self.driver.execute_script("""
                    // Get elements for To Be Executed
                    const v = document.getElementById('lbltobeexecuted');
                    const u = document.getElementById('lbltobeexecutedunit');
                    
                    // Get elements for Already Executed
                    const executedQtyElem = document.getElementById('lblexecuted');
                    const executedUnitElem = document.getElementById('lblexecutedunit');
                    
                    // Get Table Qty (Grand Qty)
                    const grandQtyElem = document.getElementById('lblGrand_Qty');
                    
                    if (!v || !u || !executedQtyElem || !executedUnitElem || !grandQtyElem) {
                        alert('Required fields not found. Contact @mrgargsir.');
                        return;
                    }
                    
                    const dn = parseFloat(v.textContent);
                    const executedQty = parseFloat(executedQtyElem.textContent);
                    const grandQty = parseFloat(grandQtyElem.textContent);
                    
                    if (isNaN(dn)) { alert('To Be Executed value is not a number.'); return; }
                    if (isNaN(executedQty)) { alert('Executed value is not a number.'); return; }
                    if (isNaN(grandQty)) { alert('Table Qty (Grand Qty) is not a number.'); return; }
                    
                    const unit = u.textContent.trim();
                    const executedUnit = executedUnitElem.textContent.trim();
                    
                    if (unit !== executedUnit) {
                        alert(`Unit mismatch! To Be Executed: ${unit}, Executed: ${executedUnit}`);
                        return;
                    }
                    
                    const al = (dn * 0.25).toFixed(3);
                    const total = (dn * 1.25).toFixed(3);
                    const remainingQty = (parseFloat(total) - executedQty).toFixed(3);
                    const pendingQty = (parseFloat(remainingQty) - grandQty).toFixed(3);
                    
                    const msg = `
                    🚨 **Extra Quantity Alert** 🚨
                ⚖️ DNIT QTY          = ${dn} ${unit}
                ➕ ALLOWANCE (25%)  = ${al} ${unit}
                ✅ TOTAL            = ${total} ${unit}
                ➖ EXECUTED QTY    = ${executedQty} ${unit}
                📌 REMAINING QTY   = ${remainingQty} ${unit}
                ➖ TABLE QTY       = ${grandQty} ${unit}
                🔄 PENDING QTY     = ${pendingQty} ${unit}
                    `;
                    
                    try {
                        alert(msg);
                    } catch (err) {
                        if (Notification.permission === 'granted') {
                            new Notification(msg);
                        } else if (Notification.permission !== 'denied') {
                            Notification.requestPermission().then(p => {
                                p === 'granted' ? new Notification(msg) : alert(msg);
                            });
                        } else {
                            alert(msg);
                        }
                    }
                """)
                print("[SUMMARY] Summary message appended.")
                return True
            else:
                print(f"[SUMMARY] Modal heading is not 'Error', got '{heading.text.strip()}'")
        except Exception as e:
            print(f"[SUMMARY] No quantity error detected or error modal not found. ({e})")
            # print(self.driver.page_source)  # Uncomment for debugging
        return False

    def process_data(self):
        """Main workflow with window management and 10-row batching"""
        print("[PROCESS] Starting main process workflow...")
        try:
            self.load_selected_excel_data_no_headers()
            self.ensure_window_visible()
            batch_size = 10
            total_rows = len(self.excel_rows)
            print(f"[PROCESS] Total rows to process: {total_rows}")
            for start in range(0, total_rows, batch_size):
                batch = self.excel_rows[start:start+batch_size]
                print(f"[PROCESS] Processing batch {start//batch_size+1} (rows {start+1}-{start+len(batch)})")
                for i, row_data in enumerate(batch):
                    portal_row_index = i + 2
                    print(f"[PROCESS] Filling portal row {portal_row_index} in batch...")
                    self.fill_portal_row(row_data, portal_row_index)
                print(f"[PROCESS] Batch {start//batch_size+1} loaded from Excel and filled in portal.")
                self.submit_data()
                print("[PROCESS] Data submitted to portal.")
                self.ensure_window_visible()
                if self.handle_quantity_error_and_summary():
                    print("[PROCESS] Quantity error handled. Stopping process.")
                    return False
                self.handle_confirmation_and_scrolling()
                print("[PROCESS] Confirmation and scrolling handled.")
                self.clear_portal_fields()
            print("[PROCESS] All batches processed successfully.")
            return True
        except Exception as e:
            print(f"[PROCESS] Processing failed: {str(e)}")
            return False
        finally:
            print("[PROCESS] Cleaning up resources...")
            self.close()

    def close(self):
        """Clean up resources"""
        print("[CLOSE] Closing resources...")
        if hasattr(self, 'driver'):
            try:
                # self.driver.quit() # skipping driver close to keep the browser open for debugging
                print("[CLOSE] WebDriver closed.")
            except Exception as e:
                print(f"[CLOSE] Error closing WebDriver: {e}")

        if hasattr(self, 'notification_popup') and self.notification_popup is not None:
            try:
                self.notification_popup.destroy()
                self.notification_popup = None  # Prevent double-destroy
                print("[CLOSE] Notification popup destroyed.")
            except Exception as e:
                print(f"[CLOSE] Error destroying notification popup: {e}")
                
        if self.notification_root:
            try:
                self.notification_root.destroy()
                self.notification_root = None  # Prevent double-destroy
                print("[CLOSE] Notification window destroyed.")
            except Exception as e:
                print(f"[CLOSE] Error destroying notification window: {e}")

# Example usage (call this before process_data):
# self.load_excel_data(self.excel_file_path)
def run_automation():
    """Run the automation process"""
    print("[RUN] Starting automation...")
    writter = HEWPwritter()
    try:
        writter.process_data()
    finally:
        writter.close()
    print("[RUN] Automation finished.")

if __name__ == "__main__":
    run_automation()
    sys.exit(0)  # Ensures the script exits cleanly