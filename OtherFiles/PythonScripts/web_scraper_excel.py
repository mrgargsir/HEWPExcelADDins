import os
import sys
import time
import subprocess
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
import os
import time

class HaryanaEBillingScraper:
    def __init__(self, website_url):
        self.website_url = "https://works.haryana.gov.in/HEWP_Login/login.aspx"
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
                    time.sleep(5)
                    return True
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
                return False
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
                
                # Wait for page to update
                time.sleep(1.5)
                
                # Verify selection took effect
                if dropdown.first_selected_option.get_attribute("value") == value:
                    # Wait for dependent elements to load
                    self.wait.until(
                        EC.presence_of_element_located((By.ID, "lbltobeexecuted"))
                    )
                    return True
                
            except Exception as e:
                print(f"⚠ Attempt {attempt+1} failed for {dropdown_id}={value}: {str(e)}")
                time.sleep(2)
        
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
            table = self.wait.until(
                EC.presence_of_element_located((By.ID, "GV_Add_to_List_POP"))
            )
            rows = table.find_elements(By.TAG_NAME, "tr")
            
            # Skip header and footer rows
            data_rows = [row for row in rows[1:-1] if "total-row" not in row.get_attribute("class")]
            
            table_data = []
            for row in data_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 11:  # Ensure we have enough columns
                    table_data.append({
                        'Sr_No': cells[0].text.strip(),
                        'Description': cells[1].text.strip(),
                        'Rate_Type': cells[2].text.strip(),
                        'Rate': cells[3].text.strip(),
                        'Unit': cells[4].text.strip(),
                        'Number': cells[5].text.strip(),
                        'Length': cells[6].text.strip(),
                        'Breadth': cells[7].text.strip(),
                        'Depth': cells[8].text.strip(),
                        'Qty': cells[9].text.strip(),
                        'Total_Quantity': cells[10].text.strip()
                    })
            return table_data
        except Exception as e:
            print(f"⚠ Error extracting table data: {str(e)}")
            return []

    def scrape_all_combinations(self):
        """Main scraping logic to process all dropdown combinations"""
        print("\nStarting data extraction process...")
        
        # Get all main head options
        main_heads = self.get_dropdown_options("ddlcomp")
        print(f"Found {len(main_heads)} Main Head options")
        
        for main_head in main_heads:
            print(f"\nProcessing Main Head: {main_head['text']}")
            
            if not self.safe_select_dropdown("ddlcomp", main_head['value']):
                continue
                
            # Get sub heads for current main head
            sub_heads = self.get_dropdown_options("ddlsubhead")
            print(f"├─ Found {len(sub_heads)} Sub Head options")
            
            for sub_head in sub_heads:
                print(f"│  ├─ Processing Sub Head: {sub_head['text']}")
                
                if not self.safe_select_dropdown("ddlsubhead", sub_head['value']):
                    continue
                    
                # Get items for current sub head
                items = self.get_dropdown_options("ddlitemnumber")
                print(f"│  │  ├─ Found {len(items)} Item options")
                
                for item in items:
                    print(f"│  │  │  ├─ Item: {item['text']}", end=" ")
                    
                    if not self.safe_select_dropdown("ddlitemnumber", item['value']):
                        continue
                        
                    # Extract data for this combination
                    quantity = self.extract_quantity()
                    table_data = self.extract_table_data()
                    print(f"[{len(table_data)} rows]")
                    
                    # Combine all data with dropdown info
                    for row in table_data:
                        self.all_data.append({
                            'Main_Head_Value': main_head['value'],
                            'Main_Head_Text': main_head['text'],
                            'Sub_Head_Value': sub_head['value'],
                            'Sub_Head_Text': sub_head['text'],
                            'Item_Number_Value': item['value'],
                            'Item_Number_Text': item['text'],
                            'Quantity_To_Execute': quantity,
                            **row
                        })
        
        print(f"\n✔ Completed! Extracted {len(self.all_data)} total records")

    def save_to_excel(self, filename="haryana_ebilling_data.xlsx"):
        """Save all collected data to Excel file with permission handling"""
        if not self.all_data:
            print("No data to save")
            return

        # Create a backup filename if primary fails
        backup_filename = None
        if os.path.exists(filename):
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            backup_filename = f"haryana_ebilling_data_{timestamp}.xlsx"

        df = pd.DataFrame(self.all_data)

        # Reorder columns for better organization
        dropdown_cols = [
            'Main_Head_Value', 'Main_Head_Text',
            'Sub_Head_Value', 'Sub_Head_Text',
            'Item_Number_Value', 'Item_Number_Text',
            'Quantity_To_Execute'
        ]
        other_cols = [col for col in df.columns if col not in dropdown_cols]

        max_attempts = 3
        attempt = 0
        success = False

        while attempt < max_attempts and not success:
            attempt += 1
            try:
                # Try saving to the primary filename first
                save_path = filename

                # If file exists and we're not on first attempt, use backup name
                if attempt > 1 and os.path.exists(filename):
                    if backup_filename:
                        save_path = backup_filename
                    else:
                        timestamp = time.strftime("%Y%m%d_%H%M%S")
                        save_path = f"haryana_ebilling_data_{timestamp}.xlsx"

                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df[dropdown_cols + other_cols].to_excel(writer, index=False)

                    # Auto-adjust column widths
                    worksheet = writer.sheets['Sheet1']
                    for column in worksheet.columns:
                        max_length = max(len(str(cell.value)) for cell in column)
                        worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

                print(f"✔ Data successfully saved to {save_path}")
                success = True

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
                        except Exception as fallback_e:
                            print(f"✖ Failed to save to fallback location: {str(fallback_e)}")
                            print("Please check your permissions or try a different directory.")
            except Exception as e:
                print(f"✖ Unexpected error saving Excel file: {str(e)}")
                break

def main():
    print("Haryana E-Billing Data Extractor")
    print("=" * 50)
    
    # Ask user for save location
    save_dir = input("Enter directory to save Excel file (leave blank for current directory): ").strip()
    if save_dir and os.path.isdir(save_dir):
        filename = os.path.join(save_dir, "haryana_ebilling_data.xlsx")
    else:
        filename = "haryana_ebilling_data.xlsx"
        print("Using current directory for output file")
    
    scraper = HaryanaEBillingScraper(
        "https://works.haryana.gov.in/E-Billing/Est_Add_Items_emb.aspx#"
    )
    
    try:
        if not scraper._check_prerequisites():
            return
        
        if not scraper.connect_to_browser():
            return
            
        scraper.ensure_window_visible()
        scraper.scrape_all_combinations()
        
        # Pass the filename to save_to_excel
        scraper.save_to_excel(filename)
        
        input("\nPress Enter to exit...")
    except Exception as e:
        print(f"\n✖ Critical error: {str(e)}")
        input("Press Enter to close...")
    finally:
        scraper.close()

if __name__ == "__main__":
    main()