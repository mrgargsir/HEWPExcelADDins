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
from tkinter import Toplevel, Label, messagebox, simpledialog
import pygetwindow as gw

class HEWPUploader:
    def __init__(self):
        self.pdf_path = r"C:\MRGARGSIR\annex.pdf"
        self.notification_root = None
        self._chrome_was_launched = False

        print("[INIT] Hiding console window")
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

        # 2. Check PDF file exists
        if not os.path.isfile(self.pdf_path):
            print(f"❌ PDF file not found: {self.pdf_path}")
            messagebox.showerror("Missing PDF", f"PDF file not found:\n{self.pdf_path}")
            sys.exit(1)
        else:
            print(f"✅ PDF file exists: {self.pdf_path}")

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

    
    

    def upload_pdf(self):
        """Selenium-only file upload approach"""
        print("[UPLOAD] Starting PDF upload process")
        self.ensure_window_visible()
        try:
            print("[UPLOAD] Waiting for file input element")
            file_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "FileUpload3"))
            )
            print(f"[UPLOAD] Sending file path: {self.pdf_path}")
            file_input.send_keys(os.path.abspath(self.pdf_path))
            print("[UPLOAD] Waiting for upload button")
            upload_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btn_add_Description"))
            )
            print("[UPLOAD] Clicking upload button")
            upload_button.click()
            time.sleep(1)
            print("[UPLOAD] PDF upload completed")
        except Exception as e:
            print(f"[UPLOAD] File upload failed: {str(e)}")
            messagebox.showerror("Upload Error", f"File upload failed: {str(e)}")
            raise

    def handle_confirmation_and_scrolling(self):
        """EXACT implementation as you specified"""
        print("[CONFIRM] Handling confirmation dialogs and scrolling")
        self.ensure_window_visible()
        try:
            # 1. First confirm any Chrome alert with Enter
            try:
                print("[CONFIRM] Checking for browser alert")
                alert = self.wait.until(lambda d: d.switch_to.alert)
                alert.accept()
                print("[CONFIRM] Alert accepted")
                time.sleep(0.5)
            except:
                print("[CONFIRM] No alert found")
                pass
            # 2. Execute your exact JavaScript sequence
            print("[CONFIRM] Executing JavaScript for UI actions")
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
            print("[CONFIRM] JavaScript executed")
            time.sleep(0.5)
        except Exception as e:
            print(f"Warning during confirmation handling: {str(e)}")

    def process_item(self):
        """Main workflow with window management"""
        print("[PROCESS] Starting main workflow")
        try:
            self.upload_pdf()
            self.handle_confirmation_and_scrolling()
            print("[PROCESS] Upload and confirmation steps complete")
            
            return True
        except Exception as e:
            print(f"[PROCESS] Processing failed: {str(e)}")
            
            return False
        finally:
            print("[PROCESS] Cleaning up resources")
            self.close()

    def close(self):
        """Clean up resources"""
        print("[CLOSE] Cleaning up")
        if hasattr(self, 'driver'):
            #self.driver.quit()        #skipping driver close to keep the browser open for debugging
            print("Browser session kept open for debugging")
        if self.notification_root:
            try:
                self.notification_root.destroy()
            except:
                pass

def run_automation():
    """Run the automation process"""
    print("[RUN] Starting automation")
    uploader = HEWPUploader()
    try:
        uploader.process_item()
    finally:
        uploader.close()

if __name__ == "__main__":
    run_automation()
    print("[EXIT] Script finished")
    sys.exit(0)  # Ensures the script exits cleanly