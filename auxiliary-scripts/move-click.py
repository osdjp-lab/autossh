import win32gui
import win32con
import win32process
import win32api
import pyautogui
import time
import tkinter as tk
from tkinter import filedialog, messagebox

TARGET = "Softship LINE"
# TARGET = "Warning"

# Add a small delay to allow pyautogui to work reliably
pyautogui.PAUSE = 0.1


def find_window():
    matches = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if TARGET.lower() in title.lower():
                matches.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    return matches[0] if matches else None


def bring_to_front(hwnd):
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    win32gui.SetForegroundWindow(hwnd)

    this_thread = win32api.GetCurrentThreadId()
    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]

    win32process.AttachThreadInput(this_thread, target_thread, True)
    win32gui.SetFocus(hwnd)
    win32process.AttachThreadInput(this_thread, target_thread, False)


def get_input_file():
    """Open a dialog for the user to select the coordinates file."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_path = filedialog.askopenfilename(
        filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")],
        title="Select coordinates file to play back..."
    )
    root.destroy()
    return file_path


def get_settings():
    """Get playback settings from the user."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # Create a simple dialog for settings
    settings_window = tk.Toplevel(root)
    settings_window.title("Playback Settings")
    settings_window.geometry("400x200")
    settings_window.attributes('-topmost', True)
    
    settings = {}
    
    # Delay between clicks
    tk.Label(settings_window, text="Delay between clicks (seconds):").pack(pady=5)
    delay_var = tk.DoubleVar(value=0.5)
    delay_spinbox = tk.Spinbox(settings_window, from_=0.1, to=5.0, increment=0.1, textvariable=delay_var, width=10)
    delay_spinbox.pack(pady=5)
    
    # Initial delay before starting
    tk.Label(settings_window, text="Initial delay before starting (seconds):").pack(pady=5)
    initial_var = tk.DoubleVar(value=2.0)
    initial_spinbox = tk.Spinbox(settings_window, from_=0.5, to=10.0, increment=0.5, textvariable=initial_var, width=10)
    initial_spinbox.pack(pady=5)
    
    # Move speed (pixels per second, 0 = instant)
    tk.Label(settings_window, text="Mouse movement speed (0 = instant, higher = slower):").pack(pady=5)
    speed_var = tk.DoubleVar(value=0)
    speed_spinbox = tk.Spinbox(settings_window, from_=0, to=1000, increment=100, textvariable=speed_var, width=10)
    speed_spinbox.pack(pady=5)
    
    def on_ok():
        settings['delay'] = delay_var.get()
        settings['initial_delay'] = initial_var.get()
        settings['speed'] = speed_var.get()
        settings_window.destroy()
        root.destroy()
    
    ok_button = tk.Button(settings_window, text="OK", command=on_ok)
    ok_button.pack(pady=10)
    
    settings_window.wait_window()
    return settings


def read_coordinates(file_path):
    """Read coordinates from a file."""
    coordinates = []
    try:
        with open(file_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        x, y = map(int, line.split(','))
                        coordinates.append((x, y))
                    except ValueError:
                        print(f"Warning: Could not parse line: {line}")
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    
    return coordinates


def playback_clicks(target_hwnd, coordinates, delay=0.5, initial_delay=2.0, speed=0):
    """Playback clicks on the target window."""
    
    # Get target window position
    rect = win32gui.GetWindowRect(target_hwnd)
    target_left, target_top, target_right, target_bottom = rect
    
    print(f"Target window position: ({target_left}, {target_top})")
    print(f"Number of clicks to perform: {len(coordinates)}")
    print(f"Initial delay: {initial_delay} seconds")
    print(f"Delay between clicks: {delay} seconds")
    
    if speed > 0:
        print(f"Mouse movement speed: {speed} pixels/second")
    else:
        print("Mouse movement: instant")
    
    # Wait for initial delay
    print(f"\nStarting playback in {initial_delay} seconds...")
    time.sleep(initial_delay)
    
    print("Starting clicks...\n")
    
    for i, (rel_x, rel_y) in enumerate(coordinates, 1):
        # Convert relative coordinates to absolute
        abs_x = target_left + rel_x
        abs_y = target_top + rel_y
        
        print(f"Click {i}/{len(coordinates)}: Moving to ({rel_x}, {rel_y}) [absolute: ({abs_x}, {abs_y})]")
        
        # Move mouse to position
        if speed > 0:
            # Calculate duration based on distance and speed
            current_x, current_y = pyautogui.position()
            distance = ((abs_x - current_x) ** 2 + (abs_y - current_y) ** 2) ** 0.5
            duration = distance / speed
            pyautogui.moveTo(abs_x, abs_y, duration=duration)
        else:
            # Move instantly
            pyautogui.moveTo(abs_x, abs_y)
        
        # Click
        pyautogui.click()
        print(f"  ✓ Clicked")
        
        # Wait before next click
        if i < len(coordinates):
            time.sleep(delay)
    
    print("\nPlayback complete!")


if __name__ == "__main__":
    # Get input file
    input_file = get_input_file()
    if not input_file:
        print("No file selected. Exiting.")
    else:
        print(f"Using coordinates file: {input_file}")
        
        # Read coordinates
        coordinates = read_coordinates(input_file)
        if not coordinates:
            print("No coordinates found in file. Exiting.")
        else:
            print(f"Loaded {len(coordinates)} coordinates")
            
            # Get playback settings
            settings = get_settings()
            
            # Find and bring target window to front
            hwnd = find_window()
            if hwnd:
                print(f"\nFound window: {hwnd:#010x}")
                bring_to_front(hwnd)
                time.sleep(0.5)
                
                # Playback clicks
                playback_clicks(
                    hwnd,
                    coordinates,
                    delay=settings['delay'],
                    initial_delay=settings['initial_delay'],
                    speed=settings['speed']
                )
            else:
                print("No window containing", TARGET, "was found.")
