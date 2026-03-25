import win32gui
import win32con
import win32process
import win32api
import time

TARGET = "Softship LINE"


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
    # Restore if minimized
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # Bring to foreground – this is allowed
    win32gui.SetForegroundWindow(hwnd)

    # -------------------------------------------------
    # Attach our thread to the target window’s thread
    # -------------------------------------------------
    # 1️⃣ ID of the **current Python thread**
    this_thread = win32api.GetCurrentThreadId()

    # 2️⃣ ID of the **target window’s thread**
    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]

    # Attach the two input queues so we can set focus
    win32process.AttachThreadInput(this_thread, target_thread, True)

    # Now we can safely give the window keyboard focus
    win32gui.SetFocus(hwnd)

    # Detach again (optional, keeps the system tidy)
    win32process.AttachThreadInput(this_thread, target_thread, False)


if __name__ == "__main__":
    hwnd = find_window()
    if hwnd:
        print(f"Found window: {hwnd:#010x}")
        bring_to_front(hwnd)
        time.sleep(0.5)
    else:
        print("No window containing", TARGET, "was found.")
