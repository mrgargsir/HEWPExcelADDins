# MRGARGSIR Chrome Debug Cleaner
# This script is designed to close all Chrome processes, kill any process using the debug port 9222,
import subprocess
import sys
import time

def is_process_running(process_name):
    """Check if a process is running."""
    result = subprocess.run(['tasklist'], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
    return process_name.lower() in result.stdout.lower()

#def kill_chrome_processes():
   # """Kill all chrome.exe processes if running."""
   # print("[CLOSE] Checking for running Chrome processes...")
   # if is_process_running('chrome.exe'):
   #     print("[CLOSE] Killing all Chrome processes...")
    #    try:
    #        subprocess.run(['taskkill', '/IM', 'chrome.exe', '/F'], check=False, creationflags=subprocess.CREATE_NO_WINDOW)
     #       print("[CLOSE] All Chrome processes killed.")
     #   except Exception as e:
     #       print(f"[CLOSE][ERROR] Error killing Chrome: {e}")
   # else:
     #   print("[CLOSE] No Chrome processes found.")

def kill_debug_port_process():
    """Find and kill all processes using port 9222."""
    print("[CLOSE] Checking for processes using debug port 9222...")
    try:
        result = subprocess.run(['netstat', '-ano'], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        killed_pids = set()
        for line in result.stdout.splitlines():
            if ":9222" in line and "LISTENING" in line:
                parts = line.split()
                pid = parts[-1]
                if pid not in killed_pids:
                    print(f"[CLOSE] Found process on port 9222 (PID {pid}), killing...")
                    try:
                        subprocess.run(['taskkill', '/PID', pid, '/F'], check=False, creationflags=subprocess.CREATE_NO_WINDOW)
                        print(f"[CLOSE] Killed process PID {pid} using port 9222.")
                        killed_pids.add(pid)
                    except Exception as e:
                        print(f"[CLOSE][ERROR] Error killing PID {pid}: {e}")
        if not killed_pids:
            print("[CLOSE] No processes found using port 9222.")
    except Exception as e:
        print(f"[CLOSE][ERROR] Error checking debug port processes: {e}")

def main():
    """Main function to clean up Chrome and debug resources."""
    print("="*40)
    print("MRGARGSIR Chrome Debug Cleaner")
    print("="*40)
    #kill_chrome_processes()
    time.sleep(1)
    kill_debug_port_process()
    print("[CLOSE] Cleanup complete.")
    print("="*40)

if __name__ == "__main__":
    try:
        main()
        sys.exit(0)
    except Exception as e:
        print(f"[CLOSE][FATAL] Unexpected error: {e}")
        sys.exit(1)