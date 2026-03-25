import win32gui

def enum_windows():
    windows = []

    def callback(hwnd, _):
        # Skip invisible windows
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                windows.append((hwnd, title))
        return True

    win32gui.EnumWindows(callback, None)
    return windows

if __name__ == "__main__":
    for hwnd, title in enum_windows():
        print(f"{hwnd:#010x}: {title}")
