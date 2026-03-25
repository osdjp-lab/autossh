import win32gui
import win32con
import win32process
import win32api
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog
import threading
import time

# TARGET = "Softship LINE"
# TARGET = "Warning"
TARGET = "Question"
# TARGET = "Required Value Missing"

output_file_path = None
overlay_hwnd = None
stop_listening = False


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


def get_output_file():
    """Open a dialog for the user to select/create an output file."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")],
        title="Save click coordinates to..."
    )
    root.destroy()
    return file_path


def create_overlay_window(target_hwnd):
    """Create a transparent overlay window."""
    global overlay_hwnd
    
    # Get target window dimensions
    rect = win32gui.GetWindowRect(target_hwnd)
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top

    # Create a simple tkinter overlay window
    overlay = tk.Tk()
    overlay.geometry(f"{width}x{height}+{left}+{top}")
    overlay.attributes('-alpha', 0.2)  # 20% opacity
    overlay.attributes('-topmost', True)
    overlay.attributes('-toolwindow', True)
    
    # Make window transparent to mouse events by binding to clicks
    overlay.configure(bg='gray')
    
    overlay_hwnd = overlay.winfo_id()
    
    print(f"Overlay created with handle: {overlay_hwnd:#010x}")

    return overlay


def listen_for_clicks(overlay, target_hwnd, output_file):
    """Listen for mouse clicks on the overlay."""
    global stop_listening, output_file_path
    
    output_file_path = output_file
    
    # Get target window position
    rect = win32gui.GetWindowRect(target_hwnd)
    target_left, target_top, target_right, target_bottom = rect
    target_width = target_right - target_left
    target_height = target_bottom - target_top

    def on_click(event):
        """Handle click events."""
        # Get absolute mouse position
        x = overlay.winfo_pointerx()
        y = overlay.winfo_pointery()
        
        # Calculate relative coordinates within target window
        rel_x = x - target_left
        rel_y = y - target_top
        
        # Ensure coordinates are within bounds
        if 0 <= rel_x <= target_width and 0 <= rel_y <= target_height:
            try:
                with open(output_file_path, 'a') as f:
                    f.write(f"{rel_x},{rel_y}\n")
                print(f"Recorded click at relative coordinates: ({rel_x}, {rel_y})")
            except Exception as e:
                print(f"Error writing to file: {e}")

    # Bind mouse button press to the overlay
    overlay.bind("<Button-1>", on_click)
    
    # Keep the overlay window responsive
    while not stop_listening:
        try:
            overlay.update()
        except:
            break
        time.sleep(0.01)


if __name__ == "__main__":
    hwnd = find_window()
    if hwnd:
        print(f"Found window: {hwnd:#010x}")
        bring_to_front(hwnd)
        time.sleep(0.5)

        # Get output file from user
        output_file = get_output_file()
        if not output_file:
            print("No output file selected. Exiting.")
        else:
            print(f"Clicks will be saved to: {output_file}")
            
            # Create overlay
            overlay = create_overlay_window(hwnd)
            print("Overlay created. Click on the overlay to record coordinates.")
            print("Close the overlay window to stop recording.")

            # Listen for clicks in a separate thread
            click_thread = threading.Thread(target=listen_for_clicks, args=(overlay, hwnd, output_file), daemon=True)
            click_thread.start()

            # Run the overlay window main loop
            try:
                overlay.mainloop()
            except KeyboardInterrupt:
                pass
            finally:
                stop_listening = True
                print("Overlay closed.")

    else:
        print("No window containing", TARGET, "was found.")
