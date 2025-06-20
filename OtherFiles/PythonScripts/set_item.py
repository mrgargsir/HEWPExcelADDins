import os
import sys
import ctypes  # NEW for window management
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pyperclip
import tkinter as tk
from tkinter import Toplevel, Label, messagebox, simpledialog
import pygetwindow as gw  # NEW for window management

class HEWPUploader:
    def __init__(self):
        self.excel_file_path = r"C:\MRGARGSIR\Length_Breadth_Depth.xlsx"
        self.notification_root = None
        self._chrome_was_launched = False  # Track if we launched Chrome

        # NEW: Hide console initially (will show if needed)
        self._hide_console()


        # Run checks (console will show if there are issues)
        if self._check_prerequisites():
            self.connect_to_browser()
            self._hide_console()  # Hide again after successful checks

    def _show_console(self):
        """Show console window"""
        if sys.platform == 'win32':
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 1)

    def _hide_console(self):
        """Hide console window"""
        if sys.platform == 'win32':
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    def _check_prerequisites(self):
        """Verify all requirements are met"""
        # NEW: Show console only if we need to display messages
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
            
        return True  # Return True if checks passed

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
                        "https://works.haryana.gov.in/HEWP_Login/login.aspx"
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
        sys.exit(1)

    def connect_to_browser(self):
        """Connect to existing Chrome browser instance"""
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
                print("After login, rerun this after all")
                sys.exit(1)
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
            # Check if the dropdown exists first
            try:
                dropdown = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#ddlitemnumber"))
                )
            except Exception:
                messagebox.showinfo(
                    "Not Ready",
                    "Item dropdown (ddlitemnumber) not found on page.\n\nPlease reach the destination page on website."
                )
                raise RuntimeError("ddlitemnumber not found on page.")

            # Continue with JS for search box
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
            
            # Now select the item
            for option in dropdown.find_elements(By.TAG_NAME, "option"):
                if item_number in option.text:
                    option.click()
                    break
            else:
                raise ValueError(f"Item '{item_number}' not found")
            
            time.sleep(1)
        except RuntimeError as e:
            # End automation if ddlitemnumber is not found
            messagebox.showerror("Automation Stopped", str(e))
            self.close()
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("Selection Error", f"Could not select item: {str(e)}")
            raise

    

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
    sys.exit(0)  # Ensures the script exits cleanly