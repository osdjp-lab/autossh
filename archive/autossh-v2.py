import pyautogui
import pygetwindow as gw
import time
import csv
import os
import sys
from pathlib import Path
from pynput import keyboard
from tkinter import Tk, filedialog
import threading

# ========================================
# CLICK LOCATION CONSTANTS
# ========================================

# Open Job
CLICK_SELECT_BUTTON = (284, 62)
CLICK_BOOKING_INPUT_FIELD = (310, 221)
CLICK_EXECUTE_QUERY = (284, 62)
CLICK_UPDATE_BUTTON = (384, 65)
CLICK_QUESTION_POPUP_OK = (332, 279)
CLICK_WARNING_POPUP_OK = (418, 147)

# Enter Costs
CLICK_COSTS_TAB = (835, 148)
CLICK_ADD_COST_LINE = (384, 65)
CLICK_REQUIRED_VALUE_OK = (378, 142)
CLICK_FIRST_COST_FIELD = (282, 325)
CLICK_DELETE_COST_LINE = (485, 64)
CLICK_CLOSE_BUTTON = (35, 64)
CLICK_QUESTION_NO_BUTTON = (101, 146)

# Save and Close
CLICK_SAVE_JOB = (235, 63)

# ========================================
# CONFIGURATION
# ========================================

SOFTSHIP_WINDOW_TITLE = "Softship LINE"
TIMEOUT = 10
ESCAPE_PRESSED = False

# ========================================
# FILE MANAGEMENT
# ========================================

COMPLETED_JOBS_FILE = "completed_jobs.csv"
FAILED_JOBS_FILE = "failed_jobs.csv"

completed_jobs = []
failed_jobs = []

# ========================================
# KEYBOARD LISTENER
# ========================================

def on_press(key):
    """Handle keyboard press - Escape key terminates program"""
    global ESCAPE_PRESSED
    try:
        if key == keyboard.Key.esc:
            ESCAPE_PRESSED = True
            print("\n\nEscape pressed! Saving jobs and terminating...")
            save_all_jobs()
            sys.exit(0)
    except AttributeError:
        pass

def start_keyboard_listener():
    """Start listening for Escape key in background thread"""
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    return listener

# ========================================
# FILE OPERATIONS
# ========================================

def select_csv_file():
    """Open file dialog to select CSV file"""
    root = Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Bring to front
    
    file_path = filedialog.askopenfilename(
        title="Select CSV file with booking data",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    
    root.destroy()
    
    if not file_path:
        print("No file selected. Exiting.")
        sys.exit(1)
    
    return file_path

def load_bookings_from_csv(file_path):
    """Load bookings from CSV file"""
    bookings = []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            
            # Validate required columns
            if reader.fieldnames is None or not all(col in reader.fieldnames for col in ['Job', 'Type', 'Currency', 'Cost']):
                print("Error: CSV must contain columns: Job, Type, Currency, Cost")
                sys.exit(1)
            
            for row in reader:
                bookings.append({
                    'Job': row['Job'].strip(),
                    'Type': row['Type'].strip(),
                    'Currency': row['Currency'].strip(),
                    'Cost': row['Cost'].strip()
                })
    
    except FileNotFoundError:
        print(f"Error: File not found - {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        sys.exit(1)
    
    if not bookings:
        print("Error: CSV file is empty")
        sys.exit(1)
    
    print(f"Loaded {len(bookings)} bookings from CSV")
    return bookings

def save_all_jobs():
    """Save completed and failed jobs to separate files"""
    global completed_jobs, failed_jobs
    
    # Save completed jobs
    if completed_jobs:
        try:
            with open(COMPLETED_JOBS_FILE, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Job', 'Type', 'Currency', 'Cost', 'Timestamp'])
                writer.writeheader()
                writer.writerows(completed_jobs)
            print(f"\nSaved {len(completed_jobs)} completed jobs to {COMPLETED_JOBS_FILE}")
        except Exception as e:
            print(f"Error saving completed jobs: {e}")
    
    # Save failed jobs
    if failed_jobs:
        try:
            with open(FAILED_JOBS_FILE, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Job', 'Type', 'Currency', 'Cost', 'Reason', 'Timestamp'])
                writer.writeheader()
                writer.writerows(failed_jobs)
            print(f"Saved {len(failed_jobs)} failed jobs to {FAILED_JOBS_FILE}")
        except Exception as e:
            print(f"Error saving failed jobs: {e}")

def add_completed_job(booking):
    """Add job to completed list"""
    global completed_jobs
    completed_job = booking.copy()
    completed_job['Timestamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
    completed_jobs.append(completed_job)
    print(f"✓ Completed: {booking['Job']}")

def add_failed_job(booking, reason="Unknown error"):
    """Add job to failed list"""
    global failed_jobs
    failed_job = booking.copy()
    failed_job['Reason'] = reason
    failed_job['Timestamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
    failed_jobs.append(failed_job)
    print(f"✗ Failed: {booking['Job']} - {reason}")

# ========================================
# UTILITY FUNCTIONS
# ========================================

def get_window(window_title):
    """Get window by title, return None if not found"""
    try:
        return gw.getWindowsWithTitle(window_title)[0]
    except IndexError:
        return None

def click(x, y, duration=0.3):
    """Click at coordinates with movement duration"""
    if ESCAPE_PRESSED:
        raise KeyboardInterrupt("Escape pressed")
    
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()
    time.sleep(0.3)

def send_hotkey(*keys):
    """Send keyboard hotkey combination"""
    if ESCAPE_PRESSED:
        raise KeyboardInterrupt("Escape pressed")
    
    pyautogui.hotkey(*keys)
    time.sleep(0.3)

def type_text(text, interval=0.05):
    """Type text with interval between characters"""
    if ESCAPE_PRESSED:
        raise KeyboardInterrupt("Escape pressed")
    
    pyautogui.typeString(text, interval=interval)
    time.sleep(0.3)

def press_key(key):
    """Press a single key"""
    if ESCAPE_PRESSED:
        raise KeyboardInterrupt("Escape pressed")
    
    pyautogui.press(key.lower())
    time.sleep(0.3)

def wait_for_window(window_title, timeout=TIMEOUT):
    """Wait for window to appear"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if ESCAPE_PRESSED:
            raise KeyboardInterrupt("Escape pressed")
        
        if get_window(window_title):
            return True
        time.sleep(0.5)
    return False

def activate_window(window_title):
    """Activate window by title"""
    window = get_window(window_title)
    if window:
        window.activate()
        time.sleep(0.3)

# ========================================
# JOB PROCESSING FUNCTIONS
# ========================================

def open_job(booking):
    """Open job workflow"""
    booking_number = booking['Job']
    print(f"\nOpening job: {booking_number}")
    
    try:
        # Focus Softship window
        activate_window(SOFTSHIP_WINDOW_TITLE)
        
        # Open query mode - Send Ctrl+S
        send_hotkey('ctrl', 's')
        time.sleep(0.5)
        
        # Click booking number input field
        click(*CLICK_BOOKING_INPUT_FIELD)
        
        # Enter booking number
        type_text(booking_number)
        
        # Execute query - Send Enter
        press_key('enter')
        time.sleep(0.5)
        
        # Enter booking - Double click
        pyautogui.moveTo(*CLICK_BOOKING_INPUT_FIELD, duration=0.3)
        pyautogui.doubleClick()
        time.sleep(0.5)
        
        # Wait for job to open
        print(f"Waiting for job to open...")
        time.sleep(TIMEOUT)
        
        # Handle "Question" window popup
        if wait_for_window("Question", timeout=3):
            print("Question window detected - saving as failed job")
            activate_window("Question")
            click(*CLICK_QUESTION_POPUP_OK)
            time.sleep(0.5)
            add_failed_job(booking, "Question popup during open")
            return False
        
        # Clear warning window
        if wait_for_window("Warning", timeout=2):
            activate_window("Warning")
            click(*CLICK_WARNING_POPUP_OK)
            time.sleep(0.5)
        
        return True
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error opening job: {e}")
        add_failed_job(booking, f"Error opening: {str(e)}")
        return False

def enter_costs(booking):
    """Enter costs workflow"""
    booking_number = booking['Job']
    print(f"Entering costs for job: {booking_number}")
    
    try:
        # Focus Softship window
        activate_window(SOFTSHIP_WINDOW_TITLE)
        
        # Click "Costs" tab
        click(*CLICK_COSTS_TAB)
        time.sleep(0.5)
        
        # Add new cost line
        click(*CLICK_ADD_COST_LINE)
        time.sleep(0.5)
        
        # Check for "Required Value Missing" window
        if wait_for_window("Required Value Missing", timeout=2):
            print("Required Value Missing window detected")
            activate_window("Required Value Missing")
            
            # Click on first line first field
            click(*CLICK_FIRST_COST_FIELD)
            time.sleep(0.3)
            
            # Copy contents (Ctrl+C)
            send_hotkey('ctrl', 'c')
            
            # Get clipboard content
            try:
                import subprocess
                if sys.platform == "win32":
                    import tkinter as tk
                    root = tk.Tk()
                    root.withdraw()
                    contents = root.clipboard_get()
                    root.destroy()
                else:
                    contents = subprocess.check_output(['xclip', '-selection', 'clipboard', '-o']).decode()
            except:
                contents = ""
            
            if contents.strip():
                # Cost entry failure
                print(f"Cost entry failure - contents not empty: {contents}")
                handle_cost_entry_failure(booking)
                return False
            else:
                # Clear closing warnings
                print("Clearing empty cost line...")
                handle_clear_closing_warnings(booking)
                return True
        
        return True
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error entering costs: {e}")
        add_failed_job(booking, f"Error entering costs: {str(e)}")
        return False

def handle_cost_entry_failure(booking):
    """Handle cost entry failure"""
    print("Handling cost entry failure...")
    
    try:
        # Click "Close" button
        click(*CLICK_CLOSE_BUTTON)
        time.sleep(0.5)
        
        # Clear "Required Value Missing" window
        activate_window("Required Value Missing")
        click(*CLICK_REQUIRED_VALUE_OK)
        time.sleep(0.5)
        
        # In "Question" window popup
        if wait_for_window("Question", timeout=2):
            activate_window("Question")
            click(*CLICK_QUESTION_NO_BUTTON)
            time.sleep(0.5)
        
        add_failed_job(booking, "Cost entry failure - required value missing")
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error handling cost entry failure: {e}")
        add_failed_job(booking, f"Error in failure handling: {str(e)}")

def handle_clear_closing_warnings(booking):
    """Clear closing warnings and errors"""
    print("Clearing closing warnings...")
    
    try:
        # Delete empty line
        click(*CLICK_DELETE_COST_LINE)
        time.sleep(0.5)
        
        # In "Question" window popup
        if wait_for_window("Question", timeout=2):
            activate_window("Question")
            click(*CLICK_QUESTION_NO_BUTTON)
            time.sleep(0.5)
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error clearing warnings: {e}")
        add_failed_job(booking, f"Error clearing warnings: {str(e)}")

def save_and_close_job(booking):
    """Save and close job workflow"""
    booking_number = booking['Job']
    print(f"Saving and closing job: {booking_number}")
    
    try:
        # Focus Softship window
        activate_window(SOFTSHIP_WINDOW_TITLE)
        
        # Save job
        click(*CLICK_SAVE_JOB)
        time.sleep(0.5)
        
        # Clear warning
        if wait_for_window("Warning", timeout=2):
            activate_window("Warning")
            click(*CLICK_WARNING_POPUP_OK)
            time.sleep(0.5)
        
        # Click "Close" button
        click(*CLICK_CLOSE_BUTTON)
        time.sleep(0.5)
        
        # Check for warning and question windows and clear them
        for window_title in ["Warning", "Question"]:
            if wait_for_window(window_title, timeout=1):
                activate_window(window_title)
                click(*CLICK_REQUIRED_VALUE_OK)
                time.sleep(0.3)
        
        add_completed_job(booking)
        return True
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error saving and closing job: {e}")
        add_failed_job(booking, f"Error saving/closing: {str(e)}")
        return False

def process_booking(booking):
    """Process a single booking"""
    try:
        if not open_job(booking):
            return False
        
        if not enter_costs(booking):
            return False
        
        if not save_and_close_job(booking):
            return False
        
        return True
    
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Error processing booking {booking['Job']}: {e}")
        add_failed_job(booking, f"Unexpected error: {str(e)}")
        return False

# ========================================
# MAIN EXECUTION
# ========================================

def main():
    """Main program execution"""
    global ESCAPE_PRESSED
    
    print("=" * 50)
    print("BOOKING PROCESSOR")
    print("=" * 50)
    print("Press ESC at any time to terminate and save progress\n")
    
    # Start keyboard listener
    start_keyboard_listener()
    
    # Select CSV file
    csv_file = select_csv_file()
    print(f"Selected file: {csv_file}\n")
    
    # Load bookings
    bookings = load_bookings_from_csv(csv_file)
    
    # Process each booking
    for i, booking in enumerate(bookings, 1):
        if ESCAPE_PRESSED:
            print("\nEscape pressed - terminating gracefully")
            break
        
        print(f"\n[{i}/{len(bookings)}] Processing booking...")
        process_booking(booking)
        time.sleep(1)  # Brief pause between bookings
    
    # Save all jobs
    print("\n" + "=" * 50)
    save_all_jobs()
    print("=" * 50)
    print("Program completed!")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProgram interrupted!")
        save_all_jobs()
        sys.exit(0)
    except Exception as e:
        print(f"\n\nFatal error: {e}")
        import traceback
        traceback.print_exc()
        save_all_jobs()
        sys.exit(1)
