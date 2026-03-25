import win32gui
import pyautogui

# Auxiliary functions

def find_window(title_contains: str):
    matches = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title_contains.lower() in title.lower():
                matches.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    return matches[0] if matches else None


def window_relative_to_screen(hwnd, rel_x, rel_y):
    """
    Convert coordinates relative to the OUTER WINDOW RECTANGLE
    into absolute screen coordinates.
    """
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return left + rel_x, top + rel_y

def get_pixel_in_window(title_contains: str, rel_x: int, rel_y: int):
    hwnd = find_window(title_contains)
    if not hwnd:
        raise RuntimeError(f"Window not found: {title_contains}")

    abs_x, abs_y = window_relative_to_screen(hwnd, rel_x, rel_y)
    return pyautogui.pixel(abs_x, abs_y)

def color_matches(actual, expected, tolerance=10):
    return all(abs(a - e) <= tolerance for a, e in zip(actual, expected))


def point_matches_color(title_contains: str, rel_x: int, rel_y: int, expected_rgb, tolerance=10):
    actual = get_pixel_in_window(title_contains, rel_x, rel_y)
    return color_matches(actual, expected_rgb, tolerance)

def screenshot_area_in_window(title_contains: str, rel_left: int, rel_top: int, width: int, height: int):
    hwnd = find_window(title_contains)
    if not hwnd:
        raise RuntimeError(f"Window not found: {title_contains}")

    abs_left, abs_top = window_relative_to_screen(hwnd, rel_left, rel_top)

    return pyautogui.screenshot(region=(abs_left, abs_top, width, height))

def area_contains_color(title_contains: str, rel_left: int, rel_top: int, width: int, height: int,
                        expected_rgb, tolerance=10):
    img = screenshot_area_in_window(title_contains, rel_left, rel_top, width, height)
    pixels = img.load()

    for y in range(img.height):
        for x in range(img.width):
            if color_matches(pixels[x, y], expected_rgb, tolerance):
                return True
    return False

def area_color_ratio(title_contains: str, rel_left: int, rel_top: int, width: int, height: int,
                     expected_rgb, tolerance=10):
    img = screenshot_area_in_window(title_contains, rel_left, rel_top, width, height)
    pixels = img.load()

    total = img.width * img.height
    matches = 0

    for y in range(img.height):
        for x in range(img.width):
            if color_matches(pixels[x, y], expected_rgb, tolerance):
                matches += 1

    return matches / total if total else 0.0

# Example uses

def is_warning_ok_enabled():
    # Example coordinates relative to Warning window
    # You would sample a pixel on the button face
    return point_matches_color("Warning", 380, 140, (240, 240, 240), tolerance=20)

def job_field_looks_active():
    ratio = area_color_ratio("Softship LINE", 280, 210, 120, 20, (255, 255, 255), tolerance=15)
    return ratio > 0.7

import time

def wait_for_point_color(title_contains: str, rel_x: int, rel_y: int, expected_rgb,
                         timeout=10, interval=0.2, tolerance=10):
    start = time.time()

    while time.time() - start < timeout:
        actual = get_pixel_in_window(title_contains, rel_x, rel_y)
        if color_matches(actual, expected_rgb, tolerance):
            return True
        time.sleep(interval)

    return False
