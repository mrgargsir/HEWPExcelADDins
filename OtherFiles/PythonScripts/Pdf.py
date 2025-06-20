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
        self.pdf_path = r"C:\MRGARGSIR\annex.pdf"
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
                print("After login, rerun the script")
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

   


    def upload_pdf(self):
        """Selenium-only file upload approach"""
        self.ensure_window_visible()
        try:
            # Open modal and upload file
           # modal_button = self.wait.until(
            #    EC.element_to_be_clickable((By.ID, 'FileUpload3"]'))
          # )
           # modal_button.click()
            
            file_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "FileUpload3"))
            )
            file_input.send_keys(os.path.abspath(self.pdf_path))
            # Wait for upload completion indicator (replace selector as needed)
          
            
            upload_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btn_add_Description"))
            )
            upload_button.click()
            time.sleep(1)
        except Exception as e:
            messagebox.showerror("Upload Error", f"File upload failed: {str(e)}")
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

 

    def process_item(self):
        """Main workflow with window management"""
        try:
        
            self.upload_pdf()
            self.handle_confirmation_and_scrolling()
            
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
    sys.exit(0)  # Ensures the script exits cleanly