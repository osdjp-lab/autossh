import pyautogui
import pygetwindow as gw
import time

# Configuration
BOOKING_NUMBER = ""  # To be filled from variable
SOFTSHIP_WINDOW_TITLE = "Softship LINE"
TIMEOUT = 10

def get_window(window_title):
    """Get window by title, return None if not found"""
    try:
        return gw.getWindowsWithTitle(window_title)[0]
    except IndexError:
        return None

def click(x, y, duration=0.5):
    """Click at coordinates with movement duration"""
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()
    time.sleep(0.3)

def send_keys(keys):
    """Send keyboard keys"""
    pyautogui.hotkey(*keys.split('+')) if '+' in keys else pyautogui.press(keys.lower())

def wait_for_window(window_title, timeout=TIMEOUT):
    """Wait for window to appear"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if get_window(window_title):
            return True
        time.sleep(0.5)
    return False

def open_job(booking_number):
    """Open job workflow"""
    print("Opening job...")
    
    # Focus Softship window
    softship = get_window(SOFTSHIP_WINDOW_TITLE)
    if softship:
        softship.activate()
    
    # Open query mode - Send Ctrl+S
    pyautogui.hotkey('ctrl', 's')
    time.sleep(0.5)
    
    # Alternative: Click Select button
    # click(284, 62)
    
    # Click booking number input field
    click(310, 221)
    time.sleep(0.3)
    
    # Enter booking number
    pyautogui.typeString(booking_number, interval=0.05)
    time.sleep(0.3)
    
    # Execute query - Send Enter
    pyautogui.press('enter')
    time.sleep(0.5)
    
    # Alternative: Click Select button
    # click(284, 62)
    
    # Enter booking - Double click
    pyautogui.moveTo(310, 221, duration=0.3)
    pyautogui.doubleClick()
    time.sleep(0.5)
    
    # Alternative: Click Update button
    # click(384, 65)
    
    # Wait for job to open
    print("Waiting for job to open...")
    time.sleep(TIMEOUT)
    
    # Handle "Question" window popup
    if wait_for_window("Question", timeout=3):
        print("Question window detected - saving failed job")
        save_failed_job(booking_number)
        question_window = get_window("Question")
        if question_window:
            question_window.activate()
        click(332, 279)
        time.sleep(0.5)
    
    # Clear warning window
    if wait_for_window("Warning", timeout=2):
        warning_window = get_window("Warning")
        if warning_window:
            warning_window.activate()
        click(418, 147)
        time.sleep(0.5)

def enter_costs():
    """Enter costs workflow"""
    print("Entering costs...")
    
    # Focus Softship window
    softship = get_window(SOFTSHIP_WINDOW_TITLE)
    if softship:
        softship.activate()
    
    # Click "Costs" tab
    click(835, 148)
    time.sleep(0.5)
    
    # Add new cost line
    click(384, 65)
    time.sleep(0.5)
    
    # Check for "Required Value Missing" window
    if wait_for_window("Required Value Missing", timeout=2):
        print("Required Value Missing window detected")
        req_window = get_window("Required Value Missing")
        if req_window:
            req_window.activate()
        
        # Click on first line first field
        click(282, 325)
        time.sleep(0.3)
        
        # Copy contents (Ctrl+C)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.3)
        
        # Get clipboard content
        import subprocess
        try:
            contents = subprocess.check_output(['xclip', '-selection', 'clipboard', '-o']).decode()
        except:
            contents = ""
        
        if contents.strip():
            # Cost entry failure
            print(f"Cost entry failure detected - contents: {contents}")
            handle_cost_entry_failure()
        else:
            # Clear closing warnings
            print("Clearing empty cost line...")
            handle_clear_closing_warnings()

def handle_cost_entry_failure():
    """Handle cost entry failure"""
    print("Handling cost entry failure...")
    
    # Click "Close" button
    click(35, 64)
    time.sleep(0.5)
    
    # Clear "Required Value Missing" window
    req_window = get_window("Required Value Missing")
    if req_window:
        req_window.activate()
    click(378, 142)
    time.sleep(0.5)
    
    # In "Question" window popup
    if wait_for_window("Question", timeout=2):
        question_window = get_window("Question")
        if question_window:
            question_window.activate()
        click(101, 146)
        time.sleep(0.5)
    
    save_failed_job(BOOKING_NUMBER)

def handle_clear_closing_warnings():
    """Clear closing warnings and errors"""
    print("Clearing closing warnings...")
    
    # Delete empty line
    click(485, 64)
    time.sleep(0.5)
    
    # In "Question" window popup
    if wait_for_window("Question", timeout=2):
        question_window = get_window("Question")
        if question_window:
            question_window.activate()
        click(101, 146)
        time.sleep(0.5)

def save_and_close_job():
    """Save and close job workflow"""
    print("Saving and closing job...")
    
    # Focus Softship window
    softship = get_window(SOFTSHIP_WINDOW_TITLE)
    if softship:
        softship.activate()
    
    # Save job
    click(235, 63)
    time.sleep(0.5)
    
    # Clear warning
    if wait_for_window("Warning", timeout=2):
        warning_window = get_window("Warning")
        if warning_window:
            warning_window.activate()
        click(418, 147)
        time.sleep(0.5)
    
    # Click "Close" button
    click(35, 64)
    time.sleep(0.5)
    
    # Check for warning and question windows and clear them
    for window_title in ["Warning", "Question"]:
        if wait_for_window(window_title, timeout=1):
            window = get_window(window_title)
            if window:
                window.activate()
            click(378, 142)
            time.sleep(0.3)

def save_failed_job(booking_number):
    """Save job details to failed additions log"""
    print(f"Saving failed job: {booking_number}")
    # Implement logging to file or database
    with open("failed_jobs.log", "a") as f:
        f.write(f"{booking_number} - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

# Main execution
if __name__ == "__main__":
    try:
        BOOKING_NUMBER = "9325078"  # Example booking number
        
        open_job(BOOKING_NUMBER)
        enter_costs()
        save_and_close_job()
        
        print("Job processing completed successfully!")
    except Exception as e:
        print(f"Error occurred: {e}")
        import traceback
        traceback.print_exc()
