from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import Toplevel, Label, messagebox, simpledialog
import time
import os
import pygetwindow as gw
import pyperclip

class HEWPUploader:
    def __init__(self):
        self.excel_file_path = r"C:\MRGARGSIR\Quantity.xlsx"
        self.notification_root = None
        self.connect_to_browser()
    
    def connect_to_browser(self):
        """Connect to existing Chrome browser instance"""
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)
        

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
        except Exception as e:
            print(f"Window management warning: {str(e)}")

    def show_notification(self, message, duration=3000):
        """Show a toast-style notification that actually works"""
        # Destroy any existing notification
        if self.notification_root:
            try:
                self.notification_root.destroy()
            except:
                pass

        # Create new notification
        self.notification_root = tk.Tk()
        self.notification_root.withdraw()  # Hide main window

        # Create popup window
        popup = Toplevel(self.notification_root)
        popup.wm_overrideredirect(True)

        # Calculate position (top-right corner)
        x_pos = self.notification_root.winfo_screenwidth() - 320
        popup.geometry(f"300x60+{x_pos}+50")

        # Style the popup
        popup.configure(background='#4CAF50')  # Green background
        Label(
            popup,
            text=message,
            fg='white',
            bg='#4CAF50',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=20
        ).pack()

        # Auto-close after duration
        def close_notification():
            try:
                popup.destroy()
                self.notification_root.destroy()
                self.notification_root = None
            except:
                pass

        popup.after(duration, close_notification)

        # Force the window to update and show
        popup.update_idletasks()
        popup.deiconify()
        popup.lift()

        # Start the event loop (this will block until the notification is closed)
        self.notification_root.mainloop()

    def ask_for_item_number(self):
        """Prompt user to enter item number (shows clipboard content as default)"""
        try:
            clipboard_text = pyperclip.paste()
        except Exception:
            clipboard_text = ""

        root = tk.Tk()
        root.withdraw()
        item_number = simpledialog.askstring(
            "Item Number Input",
            "Enter the item number (can be partial):",
            initialvalue=clipboard_text,
            parent=root
        )
        root.destroy()
        return item_number.strip() if item_number else None

    def search_and_select_item(self, item_number):
        """Search and select item that contains the number"""
        try:
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
            
            dropdown = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#ddlitemnumber"))
            )
            
            for option in dropdown.find_elements(By.TAG_NAME, "option"):
                if item_number in option.text:
                    option.click()
                    break
            else:
                raise ValueError(f"Item '{item_number}' not found")
            
            time.sleep(1)
        except Exception as e:
            messagebox.showerror("Selection Error", f"Could not select item: {str(e)}")
            raise

    def upload_file(self):
        """Selenium-only file upload approach"""
        self.ensure_window_visible()
        try:
            # Open modal and upload file
            modal_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-bs-target="#MyPopup"]'))
           )
            modal_button.click()
            
            file_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "FileUploadexcel"))
            )
            file_input.send_keys(os.path.abspath(self.excel_file_path))
            # Wait for upload completion indicator (replace selector as needed)
          
            
            upload_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btn_excel"))
            )
            upload_button.click()
            time.sleep(1)
        except Exception as e:
            messagebox.showerror("Upload Error", f"File upload failed: {str(e)}")
            raise

    def copy_excel_data(self):
        """Copy data from Excel"""
        self.ensure_window_visible()
        try:
            # Try both possible copy buttons
            try:
                copy_btn = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "btncopyexcel"))
                )
                if copy_btn.get_attribute("value") == "Copy Excel Data":
                    copy_btn.click()
            except:
                copy_btn = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "Button6"))
                )
                if copy_btn.get_attribute("value") == "Copy Excel Data":
                    copy_btn.click()
            
            time.sleep(0.5)
            
            # Execute the exact popup handling sequence you requested
            self.handle_confirmation_and_scrolling()
        except Exception as e:
            messagebox.showerror("Copy Error", f"Data copy failed: {str(e)}")
            raise

    def handle_confirmation_and_scrolling(self):
        """EXACT implementation as you specified"""
        self.ensure_window_visible()
        try:
            # 1. First confirm any Chrome alert with Enter
            try:
                alert = self.wait.until(lambda d: d.switch_to.alert)
                alert.accept()
                time.sleep(0.5)
            except:
                pass
            # 2. Execute your exact JavaScript sequence
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
            
            time.sleep(0.5)
        except Exception as e:
            print(f"Warning during confirmation handling: {str(e)}")

    def select_rate_type_with_script(self):
        """Select Rate Type using Selenium, skip if not present, fallback to highlighting other dropdowns."""
        from selenium.common.exceptions import NoSuchElementException, TimeoutException
        import time


        try:
            # Try to find the Rate Type dropdown, skip if not found
            try:
                rate_type_select = self.driver.find_element(By.ID, "ddlRate_Type")
            except NoSuchElementException:
                # If not present, skip this step
                print("Contracter Login, Rate_Type not found, skipping Rate Type selection.")
                return

            # Get all options
            options = rate_type_select.find_elements(By.TAG_NAME, "option")
            # Try to select in order of preference
            preferred = ["Through Rate", "Rate", "Labour Rate"]
            selected = False
            for pref in preferred:
                for opt in options:
                    if opt.text.strip() == pref:
                        opt.click()
                        selected = True
                        break
                if selected:
                    break

            # If none of the preferred, select first non-default
            if not selected:
                for opt in options:
                    if opt.get_attribute("value") != "Select One":
                        opt.click()
                        selected = True
                        break

            # Highlight the dropdown for user feedback
            self.driver.execute_script(
                "arguments[0].style.outline='3px solid orange'; arguments[0].style.backgroundColor='#bf360c';", 
                rate_type_select
            )
            time.sleep(1)
            self.driver.execute_script(
                "arguments[0].style.outline=''; arguments[0].style.backgroundColor='';", 
                rate_type_select
            )

        except Exception as e:
            print(f"Error in select_rate_type_with_script: {e}")

    def process_item(self):
        """Main workflow with window management"""
        try:
            item_number = self.ask_for_item_number()
            if not item_number:
                return
            self.search_and_select_item(item_number)
            self.select_rate_type_with_script()  # <-- Call the new function here
            self.upload_file()
            self.copy_excel_data()
            
            self.show_notification("ALL DONE!")
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            return False
        finally:
            self.close()

    def close(self):
        """Clean up resources"""
        if hasattr(self, 'driver'):
            self.driver.quit()
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