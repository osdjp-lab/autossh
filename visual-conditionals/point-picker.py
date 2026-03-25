import csv
import time
from pathlib import Path

import pyautogui
from pynput import mouse
import win32gui


########################################
# CONFIG
########################################

# Set to a partial window title if you want relative coordinates as well.
# Example: "Softship LINE"
# Set to None to disable relative coordinate output.
TARGET_WINDOW_TITLE = None

# Output CSV file
OUTPUT_FILE = Path("picked_points.csv")


########################################
# HELPERS
########################################

def find_window(title_contains: str):
    matches = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title_contains.lower() in title.lower():
                matches.append((hwnd, title))
        return True

    win32gui.EnumWindows(callback, None)
    return matches[0] if matches else None


def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def get_relative_coords(x, y, title_contains):
    result = find_window(title_contains)
    if not result:
        return None, None, None

    hwnd, title = result
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)

    rel_x = x - left
    rel_y = y - top

    return rel_x, rel_y, title


########################################
# MAIN
########################################

def main():
    print("Point picker started.")
    print("Left click = record point")
    print("Right click = stop")
    print(f"Output file: {OUTPUT_FILE.resolve()}")
    print()

    with OUTPUT_FILE.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "timestamp",
            "screen_x",
            "screen_y",
            "relative_x",
            "relative_y",
            "rgb_r",
            "rgb_g",
            "rgb_b",
            "hex",
            "target_window_title"
        ])

        def on_click(x, y, button, pressed):
            if not pressed:
                return

            if button == mouse.Button.right:
                print("Stopping point picker...")
                return False

            if button == mouse.Button.left:
                rgb = pyautogui.pixel(x, y)
                hex_color = rgb_to_hex(rgb)

                rel_x = ""
                rel_y = ""
                title = ""

                if TARGET_WINDOW_TITLE:
                    rel_x, rel_y, title = get_relative_coords(x, y, TARGET_WINDOW_TITLE)
                    rel_x = "" if rel_x is None else rel_x
                    rel_y = "" if rel_y is None else rel_y
                    title = "" if title is None else title

                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")

                writer.writerow([
                    timestamp,
                    x,
                    y,
                    rel_x,
                    rel_y,
                    rgb[0],
                    rgb[1],
                    rgb[2],
                    hex_color,
                    title
                ])
                f.flush()

                print(f"[{timestamp}]")
                print(f"Screen: ({x}, {y})")
                if TARGET_WINDOW_TITLE:
                    print(f"Relative to '{TARGET_WINDOW_TITLE}': ({rel_x}, {rel_y})")
                print(f"Color: RGB{rgb}  HEX={hex_color}")
                print("-" * 40)

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

    print(f"Saved points to: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()