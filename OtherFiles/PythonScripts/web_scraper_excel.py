import os
import sys
import time
import subprocess
import ctypes
import importlib.util
import pandas as pd
import pyautogui
import pyperclip
import pygetwindow as gw
from tkinter import messagebox
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

class HaryanaEBillingScraper:
    def __init__(self, website_url):
        self.website_url = "https://works.haryana.gov.in/HEWP_Login/login.aspx"
        self.driver = None
        self.wait = None
        self.all_data = []
        self.wait_time = 15
        self.max_retries = 3
        self._chrome_was_launched = False
        self.notification_root = None


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
            "selenium", "pyautogui", "pyperclip", "pygetwindow", "tkinter",
            "pandas", "openpyxl", "tqdm", "webdriver_manager", "re"
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

    def extract_quantity(self):
        """Extract the quantity to be executed"""
        try:
            return self.driver.find_element(By.ID, "lbltobeexecuted").text.strip()
        except:
            return "N/A"

    def extract_table_data(self):
        """Extract all data from the main table"""
        try:
            try:
                savecheck_btn = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "btn_Final_Save"))
                )
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", savecheck_btn)
            except TimeoutException:
                print("Page not loaded properly.")
                return []

            try:
                table = WebDriverWait(self.driver, 0.5).until(
                    EC.presence_of_element_located((By.ID, "GV_Add_to_List_POP"))
                )
            except TimeoutException:
                print("Table not found. Assuming all rows blanks.")
                return []
            
            # Scroll to the table to ensure it's loaded
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", table)
            time.sleep(0.5)  # Allow time for table to render
            rows = table.find_elements(By.TAG_NAME, "tr")
            
            # Skip header and footer rows
            data_rows = [row for row in rows[1:-1] if "total-row" not in row.get_attribute("class")]
            
            table_data = []
            for row in data_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 11:
                   
                    try:
                        qty = float(cells[10].text.strip().replace(',', ''))
                    except ValueError:
                        qty = cells[10].text.strip()
                    try:
                        rate = float(cells[3].text.strip().replace(',', ''))
                    except ValueError:
                        rate = cells[3].text.strip()
                    try:
                        number = float(cells[5].text.strip().replace(',', ''))
                    except ValueError:
                        number = cells[5].text.strip()
                    try:
                        length = float(cells[6].text.strip().replace(',', ''))
                    except ValueError:
                        length = cells[6].text.strip()
                    try:
                        breadth = float(cells[7].text.strip().replace(',', ''))
                    except ValueError:
                        breadth = cells[7].text.strip()
                    try:
                        depth = float(cells[8].text.strip().replace(',', ''))
                    except ValueError:
                        depth = cells[8].text.strip()
                    # Only calculate amount if both qty and rate are numbers
                    if isinstance(qty, (int, float)) and isinstance(rate, (int, float)):
                        amount = qty * rate
                    else:
                        amount = ""

                    table_data.append({
                        'Description': cells[1].text.strip(),
                        'Number': number,
                        'Length': length,
                        'Breadth': breadth,
                        'Depth': depth,
                        'Qty': qty,
                        'Unit': cells[4].text.strip(),
                        'Rate_Type': cells[2].text.strip(),
                        'Rate': rate,
                        'Amount': amount,
                        'Quantity Previously Executed': cells[9].text.strip()
                    })
            return table_data
        except Exception as e:
            print(f"⚠ Error extracting table data: {str(e)}")
            return []

    def scrape_all_combinations(self):
        """Main scraping logic to process all dropdown combinations"""
        print("\nStarting data extraction process...")

        # Get all main head options, skip empty and "Select One"
        main_heads = [
            mh for mh in self.get_dropdown_options("ddlcomp")
            if mh['text'].strip() and mh['text'].strip().lower() != "select one"
        ]
        print(f"Found {len(main_heads)} Main Head options (excluding empty/'Select One')")

        for main_head in tqdm(main_heads, desc="Main Heads"):
            print(f"\nProcessing Main Head: {main_head['text']}")

            if not self.safe_select_dropdown("ddlcomp", main_head['value']):
                continue

            # Get sub heads for current main head, skip empty and "Select One"
            sub_heads = [
                sh for sh in self.get_dropdown_options("ddlsubhead")
                if sh['text'].strip() and sh['text'].strip().lower() != "select one"
            ]
            print(f"├─ Found {len(sub_heads)} Sub Head options (excluding empty/'Select One')")

            for sub_head in sub_heads:
                if not self.safe_select_dropdown("ddlsubhead", sub_head['value']):
                    continue

                # Get items for current sub head, skip empty and "Select One"
                items = [
                    it for it in self.get_dropdown_options("ddlitemnumber")
                    if it['text'].strip() and it['text'].strip().lower() != "select one"
                ]

                for item in items:
                    if not self.safe_select_dropdown("ddlitemnumber", item['value']):
                        continue

                    
                    print(f"├─ Processing Sub Head: {sub_head['text']} - Item: {item['text']}") 
                    # Extract data for this combination
                    quantity = self.extract_quantity()
                    table_data = self.extract_table_data()

                   

                    # Combine all data with dropdown info
                    for row in table_data:
                        # Extract only the text inside [ ]
                        match = re.search(r'\[([^\]]+)\]', item['text'])
                        hsr_item_number = match.group(1).strip() if match else item['text'].strip()


                        # Try to convert quantity to float, else keep as text
                        try:
                            total_quantity = float(quantity.replace(',', ''))
                        except (ValueError, AttributeError):
                            total_quantity = quantity

                        self.all_data.append({
                            'Main_Head': main_head['text'],
                            'Sub_Head': sub_head['text'],
                            'HSR Item_Number': hsr_item_number,
                            'Total Quantity as per DNIT': total_quantity,
                            **row
                        })

        print(f"\n✔ Completed! Extracted {len(self.all_data)} total records")

    def save_to_excel(self, filename="haryana_ebilling_data.xlsx"):
        """Save all collected data to Excel file with permission handling and summary sheet"""
        if not self.all_data:
            print("No data to save")
            return None

        df = pd.DataFrame(self.all_data)

        # Reorder columns for better organization
        dropdown_cols = [
            'Main_Head',
            'Sub_Head',
            'HSR Item_Number',
            'Total Quantity as per DNIT'
        ]
        other_cols = [col for col in df.columns if col not in dropdown_cols]

        max_attempts = 3
        attempt = 0
        success = False
        save_path = filename

        while attempt < max_attempts and not success:
            attempt += 1
            try:
                if attempt > 1 and os.path.exists(filename):
                    # Create a backup filename if primary fails
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
                    backup_filename = os.path.join(desktop_dir, f"haryana_ebilling_data_{timestamp}.xlsx")

                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    # Always write main data
                    df[dropdown_cols + other_cols].to_excel(writer, index=False, sheet_name="Sheet1")

                    try:
                        # --- SUMMARY SHEET LOGIC STARTS HERE ---
                        summary_cols = [
                            "HSR Item_Number", "Qty", "Unit", "Rate", "Amount"
                        ]
                        df_summary = df.copy()
                        df_summary["Qty"] = pd.to_numeric(df_summary["Qty"], errors="coerce").fillna(0)
                        df_summary["Amount"] = pd.to_numeric(df_summary["Amount"], errors="coerce").fillna(0)

                        summary = (
                            df_summary.groupby("HSR Item_Number", as_index=False)
                            .agg({
                                "Qty": "sum",
                                "Unit": "first",
                                "Rate": "first",
                                "Amount": "sum"
                            })
                        )

                        summary.insert(0, "Sr. No.", range(1, len(summary) + 1))

                        summary_sheet = writer.book.create_sheet("Portal Summary")
                        summary_sheet.append(["Data from: Portal"])
                        summary_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
                        cell = summary_sheet.cell(row=1, column=1)
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal="center")

                        summary_headers = ["Sr. No.", "HSR Item_Number", "Qty", "Unit", "Rate", "Amount"]
                        summary_sheet.append(summary_headers)

                        for row in summary.itertuples(index=False):
                            summary_sheet.append(row)

                        total_qty = summary["Qty"].sum()
                        total_amount = summary["Amount"].sum()
                        total_row = ["Total", "", total_qty, "", "", total_amount]
                        summary_sheet.append(total_row)

                        for col in range(1, 6):
                            cell = summary_sheet.cell(row=summary_sheet.max_row, column=col)
                            cell.font = Font(bold=True)

                        # Auto-adjust column widths
                        for ws in [writer.sheets['Sheet1'], summary_sheet]:
                            for column_cells in ws.columns:
                                for cell in column_cells:
                                    if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                                        col_letter = cell.column_letter
                                        break
                                else:
                                    continue
                                max_length = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=0)
                                ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

                    except Exception as summary_e:
                        print(f"⚠ Error creating summary or formatting: {summary_e}")
                        print("Main data sheet will still be saved.")

                print(f"✔ Data successfully saved to {save_path}")
                success = True
                return save_path

            except PermissionError as e:
                print(f"⚠ Attempt {attempt} failed: {str(e)}")
                if attempt < max_attempts:
                    if os.path.exists(filename):
                        print(f"File '{filename}' may be locked. Please close it if open in Excel.")
                    print("Retrying...")
                    time.sleep(2)  # Wait before retrying
                else:
                    # Try saving to Documents folder as last resort
                    documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
                    if os.path.exists(documents_path):
                        fallback_path = os.path.join(documents_path, f"haryana_ebilling_data_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
                        try:
                            df[dropdown_cols + other_cols].to_excel(fallback_path, index=False)
                            print(f"✔ Data saved to fallback location: {fallback_path}")
                            success = True
                            return fallback_path  # Return fallback path
                        except Exception as fallback_e:
                            print(f"✖ Failed to save to fallback location: {str(fallback_e)}")
                            print("Please check your permissions or try a different directory.")
            except Exception as e:
                print(f"✖ Unexpected error saving Excel file: {str(e)}")
                break
        return None  # If all attempts fail

    def close(self):
        if self.driver:
            # self.driver.quit()
            print("skipped Closing browser session...")
            pass

# In your main function:
def main():
    print("Haryana E-Billing Data Extractor")
    print("=" * 50)
    
    # Standard Desktop path
    desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
    # OneDrive Desktop path (common for Windows)
    onedrive_dir = os.path.join(os.path.expanduser('~'), 'OneDrive', 'Desktop')

    # Prefer standard Desktop, else OneDrive Desktop, else create standard Desktop
    if os.path.exists(desktop_dir):
        chosen_dir = desktop_dir
    elif os.path.exists(onedrive_dir):
        chosen_dir = onedrive_dir
    else:
        os.makedirs(desktop_dir, exist_ok=True)
        chosen_dir = desktop_dir

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    default_filename = os.path.join(chosen_dir, f"haryana_ebilling_data_{timestamp}.xlsx")
    
    filename = default_filename
    print(f"Using Desktop for output file: {filename}")
    
    scraper = HaryanaEBillingScraper(
        "https://works.haryana.gov.in/E-Billing/Est_Add_Items_emb.aspx#"
    )
    
    try:
        print("[START] Starting the scraping process...")
            
        scraper.ensure_window_visible()
        scraper.scrape_all_combinations()
        
        # Save and get the actual saved path
        saved_path = scraper.save_to_excel(filename)
        
        # Auto-open the Excel file only if it exists
        if saved_path and os.path.exists(saved_path):
            try:
                os.startfile(saved_path)
                print(f"Opened Excel file: {saved_path}")
                sys.exit(0)  # Exit gracefully, no error
            except Exception as e:
                print(f"Could not open Excel file automatically: {e}")
        else:
            print("Excel file was not created (no data extracted).")
        
    except Exception as e:
        print(f"\n✖ Critical error: {str(e)}")
        
    finally:
        scraper.close()
if __name__ == "__main__":
    main()
