import csv
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import win32gui
import win32con
import win32process
import win32api

import pyautogui
import pyperclip
import keyboard


########################################
# PYAUTOGUI SAFETY
########################################

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


########################################
# CLICK LOCATION COORDINATES
# These are assumed to be relative to the
# OUTER WINDOW RECTANGLE of their target window.
########################################

SELECT_BTN = (284, 62)
JOB_NR_FIELD = (310, 221)
UPDATE_BTN = (384, 65)
OPEN_WITH_USER_IN_BOOKING_NO_BTN = (332, 279)
WARNING_OK = (422, 161)
COSTS_TAB = (835, 148)
INSERT_BTN = (384, 65)
SAVE_BTN = (235, 63)
CLOSE_BTN = (35, 64)
REQ_VALUE_MISSING_OK = (378, 142)
FIRST_LINE_FIRST_FIELD = (282, 325)
CLOSE_WITHOUT_SAVING_YES = (216, 162)
DELETE_BTN = (485, 64)
DELETE_YES = (101, 146)


########################################
# CONSTANTS
########################################

INTER_ACTION_WAIT = 3
WAIT_QUERY_SWITCH = 3
WAIT_TILL_JOB_OPEN = 10
WAIT_TILL_SAVE = 10

REQUIRED_COLUMNS = ["Job", "Type", "Currency", "Cost", "Per"]


########################################
# WINDOW TITLES
########################################

SOFTSHIP_WINDOW = "Softship LINE"
WARNING_WINDOW = "Warning"
QUESTION_WINDOW = "Question"
REQ_MISSING_WINDOW = "Required Value Missing"


########################################
# STOP HANDLING
########################################

STOP_REQUESTED = False


class StopExecution(Exception):
    """Raised when user presses Escape to stop the script."""
    pass


def request_stop():
    global STOP_REQUESTED
    STOP_REQUESTED = True
    print("\nESC pressed - stopping execution safely...")


def check_stop():
    if STOP_REQUESTED:
        raise StopExecution()


def interruptible_sleep(seconds, interval=0.1):
    elapsed = 0.0
    while elapsed < seconds:
        check_stop()
        sleep_for = min(interval, seconds - elapsed)
        time.sleep(sleep_for)
        elapsed += sleep_for


########################################
# HELPER FUNCTIONS
########################################

def find_window(title_contains: str):
    """
    Return the first visible top-level window handle whose title contains title_contains.
    Returns None if no matching window is found.
    """
    check_stop()
    matches = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title_contains.lower() in title.lower():
                matches.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    return matches[0] if matches else None


def window_exists(title_contains: str) -> bool:
    check_stop()
    return find_window(title_contains) is not None


def bring_to_front(hwnd):
    """
    Restore and bring the specified window handle to the foreground.
    Uses thread input attachment to improve reliability.
    """
    check_stop()

    if not hwnd:
        return

    # Restore if minimized
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        interruptible_sleep(0.2)

    # Show window
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    interruptible_sleep(0.1)

    # Get thread IDs
    foreground_hwnd = win32gui.GetForegroundWindow()
    current_thread_id = win32api.GetCurrentThreadId()
    target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
    foreground_thread_id = 0

    if foreground_hwnd:
        foreground_thread_id = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]

    attached_to_target = False
    attached_to_foreground = False

    try:
        if current_thread_id != target_thread_id:
            win32process.AttachThreadInput(current_thread_id, target_thread_id, True)
            attached_to_target = True

        if foreground_thread_id and foreground_thread_id != current_thread_id:
            win32process.AttachThreadInput(current_thread_id, foreground_thread_id, True)
            attached_to_foreground = True

        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetActiveWindow(hwnd)
        win32gui.SetFocus(hwnd)

    finally:
        if attached_to_target:
            try:
                win32process.AttachThreadInput(current_thread_id, target_thread_id, False)
            except Exception:
                pass

        if attached_to_foreground:
            try:
                win32process.AttachThreadInput(current_thread_id, foreground_thread_id, False)
            except Exception:
                pass


def activate_window(title_contains: str) -> bool:
    """
    Find a window by title fragment and bring it to the foreground.
    Returns True if successful, False otherwise.
    """
    check_stop()
    hwnd = find_window(title_contains)
    if not hwnd:
        print(f"No window found containing: {title_contains}")
        return False

    try:
        bring_to_front(hwnd)
        interruptible_sleep(0.2)
        return True
    except Exception as e:
        print(f"Failed to activate window '{title_contains}': {e}")
        return False


def get_window_rect(hwnd):
    """
    Return (left, top, right, bottom) of the outer window rectangle.
    """
    check_stop()
    return win32gui.GetWindowRect(hwnd)


def click_relative_to_window(hwnd, coord, pause=0.3):
    """
    Click using coordinates relative to the OUTER WINDOW RECTANGLE.
    coord = (x, y)
    """
    check_stop()

    if not hwnd:
        raise RuntimeError("Cannot click: invalid window handle.")

    left, top, right, bottom = get_window_rect(hwnd)
    abs_x = left + coord[0]
    abs_y = top + coord[1]

    pyautogui.click(abs_x, abs_y)
    interruptible_sleep(pause)


def click_in_window(title_contains: str, coord, pause=0.3, activate_first=True):
    """
    Find a window by title, optionally activate it, then click coordinates
    relative to that window.
    """
    check_stop()

    hwnd = find_window(title_contains)
    if not hwnd:
        raise RuntimeError(f"Window not found: {title_contains}")

    if activate_first:
        activate_window(title_contains)

        # Re-acquire handle after activation just in case
        hwnd = find_window(title_contains)
        if not hwnd:
            raise RuntimeError(f"Window disappeared after activation: {title_contains}")

    click_relative_to_window(hwnd, coord, pause=pause)


def write_text(text: str, interval: float = 0.03):
    for ch in str(text):
        check_stop()
        pyautogui.write(ch)
        interruptible_sleep(interval)


def press_tab(times=1, pause=0.2):
    for _ in range(times):
        check_stop()
        pyautogui.press("tab")
        interruptible_sleep(pause)


def save_row(writer, row, flush_file=None):
    writer.writerow(row)
    if flush_file:
        flush_file.flush()


def split_cost(cost_value):
    """
    Splits cost into integer and decimal components as strings.
    Examples:
        123.45 -> ("123", "45")
        "123,45" -> ("123", "45")
        100 -> ("100", "00")
    """
    s = str(cost_value).strip().replace(",", ".")
    if "." in s:
        integer_part, decimal_part = s.split(".", 1)
        decimal_part = (decimal_part + "00")[:2]
    else:
        integer_part = s
        decimal_part = "00"
    return integer_part, decimal_part


def type_cost_symbol_by_symbol(cost_value):
    integer_part, decimal_part = split_cost(cost_value)

    for ch in integer_part:
        check_stop()
        pyautogui.write(ch)
        interruptible_sleep(0.03)

    check_stop()
    pyautogui.write(",")
    interruptible_sleep(0.03)

    for ch in decimal_part:
        check_stop()
        pyautogui.write(ch)
        interruptible_sleep(0.03)


def copy_selected_text():
    check_stop()
    pyperclip.copy("")
    pyautogui.hotkey("ctrl", "c")
    interruptible_sleep(0.3)
    return pyperclip.paste().strip()


########################################
# CORE AUTOMATION STEPS
########################################

def open_job(job_number):
    """
    Open the job in Softship LINE.
    Returns:
        'ok' if opened successfully,
        'question_popup' if Question popup occurred,
        'warning_handled' if Warning popup was cleared.
    """
    activate_window(SOFTSHIP_WINDOW)

    # Open query mode
    click_in_window(SOFTSHIP_WINDOW, SELECT_BTN)

    # Wait till query ends
    interruptible_sleep(WAIT_QUERY_SWITCH)

    # Job number input field
    click_in_window(SOFTSHIP_WINDOW, JOB_NR_FIELD)

    # Move cursor to end of current text
    pyautogui.hotkey("ctrl", "right")

    # Clear field (as in your original logic)
    for _ in range(10):
        check_stop()
        pyautogui.press("backspace")

    interruptible_sleep(INTER_ACTION_WAIT)

    # Input job number
    write_text(job_number)

    # Execute query
    click_in_window(SOFTSHIP_WINDOW, SELECT_BTN)

    # Wait till query ends
    interruptible_sleep(WAIT_QUERY_SWITCH)

    # Enter booking / update
    click_in_window(SOFTSHIP_WINDOW, UPDATE_BTN)

    # Wait till open
    interruptible_sleep(WAIT_TILL_JOB_OPEN)

    # If "Question" window popup occurs
    if window_exists(QUESTION_WINDOW):
        click_in_window(QUESTION_WINDOW, OPEN_WITH_USER_IN_BOOKING_NO_BTN)
        return "question_popup"

    # Clear warning if present
    if window_exists(WARNING_WINDOW):
        click_in_window(WARNING_WINDOW, WARNING_OK)
        interruptible_sleep(INTER_ACTION_WAIT)
        return "warning_handled"

    return "ok"


def enter_cost(cost_type, currency, cost, per_value):
    """
    Enter cost details in Softship LINE.
    """
    activate_window(SOFTSHIP_WINDOW)

    # Enter Costs tab
    click_in_window(SOFTSHIP_WINDOW, COSTS_TAB)
    interruptible_sleep(INTER_ACTION_WAIT)

    # Add new cost line
    click_in_window(SOFTSHIP_WINDOW, INSERT_BTN)
    interruptible_sleep(INTER_ACTION_WAIT)

    # First field of added line is auto selected
    write_text(cost_type)
    press_tab()

    write_text("1")
    press_tab()

    write_text("+")
    press_tab()

    write_text(currency)
    press_tab()

    type_cost_symbol_by_symbol(cost)
    press_tab()

    write_text(per_value)


def save_changes():
    """
    Click save and handle post-save popups.
    Returns:
        'saved'
        'warning_after_save'
        'required_value_missing'
    """
    activate_window(SOFTSHIP_WINDOW)

    click_in_window(SOFTSHIP_WINDOW, SAVE_BTN)
    interruptible_sleep(WAIT_TILL_SAVE)

    if window_exists(WARNING_WINDOW):
        click_in_window(WARNING_WINDOW, WARNING_OK)
        interruptible_sleep(INTER_ACTION_WAIT)
        return "warning_after_save"

    if window_exists(REQ_MISSING_WINDOW):
        click_in_window(REQ_MISSING_WINDOW, REQ_VALUE_MISSING_OK)
        interruptible_sleep(INTER_ACTION_WAIT)
        return "required_value_missing"

    return "saved"


def handle_required_value_missing():
    """
    Handles the 'Required Value Missing' logic.
    Returns copied contents of first field after clicking it.
    """
    activate_window(SOFTSHIP_WINDOW)
    click_in_window(SOFTSHIP_WINDOW, FIRST_LINE_FIRST_FIELD)

    # Contents are auto-selected
    contents = copy_selected_text()
    return contents


def close_without_saving():
    """
    Close the job and confirm close without saving.
    """
    activate_window(SOFTSHIP_WINDOW)
    click_in_window(SOFTSHIP_WINDOW, CLOSE_BTN)

    interruptible_sleep(INTER_ACTION_WAIT)

    if window_exists(REQ_MISSING_WINDOW):
        click_in_window(REQ_MISSING_WINDOW, REQ_VALUE_MISSING_OK)
        interruptible_sleep(INTER_ACTION_WAIT)

    if window_exists(QUESTION_WINDOW):
        click_in_window(QUESTION_WINDOW, CLOSE_WITHOUT_SAVING_YES)
        interruptible_sleep(INTER_ACTION_WAIT)


def delete_empty_line_and_close():
    """
    Delete an empty cost line, confirm deletion, then close.
    """
    activate_window(SOFTSHIP_WINDOW)
    click_in_window(SOFTSHIP_WINDOW, DELETE_BTN)
    interruptible_sleep(INTER_ACTION_WAIT)

    if window_exists(QUESTION_WINDOW):
        click_in_window(QUESTION_WINDOW, DELETE_YES)
        interruptible_sleep(INTER_ACTION_WAIT)

    activate_window(SOFTSHIP_WINDOW)
    click_in_window(SOFTSHIP_WINDOW, CLOSE_BTN)
    interruptible_sleep(INTER_ACTION_WAIT)


########################################
# CSV HANDLING
########################################

def validate_csv_columns(fieldnames):
    if not fieldnames:
        raise ValueError("CSV file has no header row.")

    missing = [col for col in REQUIRED_COLUMNS if col not in fieldnames]
    if missing:
        raise ValueError(f"CSV file is missing required columns: {missing}")


def get_output_paths(input_csv_path: Path):
    failed_path = input_csv_path.with_name(f"{input_csv_path.stem}_failed.csv")
    succeeded_path = input_csv_path.with_name(f"{input_csv_path.stem}_succeeded.csv")
    return failed_path, succeeded_path


########################################
# MAIN
########################################

def main():
    global STOP_REQUESTED

    keyboard.add_hotkey("esc", request_stop)

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        input_path_str = filedialog.askopenfilename(
            title="Select input CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        root.destroy()

        # User cancelled dialog
        if not input_path_str:
            messagebox.showinfo("No file selected", "Operation cancelled - no CSV file was selected.")
            return

        check_stop()

        input_csv_path = Path(input_path_str)

        if not input_csv_path.exists():
            messagebox.showerror("File not found", f"File not found:\n{input_csv_path}")
            return

        failed_path, succeeded_path = get_output_paths(input_csv_path)

        with input_csv_path.open("r", newline="", encoding="utf-8-sig") as infile:
            reader = csv.DictReader(infile)
            validate_csv_columns(reader.fieldnames)

            fieldnames = reader.fieldnames
            all_rows = list(reader)

        with failed_path.open("w", newline="", encoding="utf-8") as failed_file, \
             succeeded_path.open("w", newline="", encoding="utf-8") as succeeded_file:

            failed_writer = csv.DictWriter(failed_file, fieldnames=fieldnames)
            succeeded_writer = csv.DictWriter(succeeded_file, fieldnames=fieldnames)

            failed_writer.writeheader()
            succeeded_writer.writeheader()

            check_stop()

            for idx, row in enumerate(all_rows):
                row_number = idx + 1
                row_already_written = False

                job = str(row["Job"]).strip()
                cost_type = str(row["Type"]).strip()
                currency = str(row["Currency"]).strip()
                cost = str(row["Cost"]).strip()
                per_value = str(row["Per"]).strip()

                print(f"\nProcessing row {row_number}: Job={job}")

                try:
                    check_stop()

                    ########################################
                    # OPEN JOB
                    ########################################
                    open_result = open_job(job)
                    print(open_result)

                    if open_result == "question_popup":
                        print(f"Job {job}: Question popup occurred while opening. Marked as failed.")
                        save_row(failed_writer, row, failed_file)
                        row_already_written = True
                        continue

                    ########################################
                    # ENTER COST
                    ########################################
                    enter_cost(cost_type, currency, cost, per_value)

                    ########################################
                    # SAVE CHANGES
                    ########################################
                    save_result = save_changes()

                    if save_result == "warning_after_save":
                        print(f"Job {job}: Warning after save. Closing and marking as failed.")
                        activate_window(SOFTSHIP_WINDOW)
                        click_in_window(SOFTSHIP_WINDOW, CLOSE_BTN)
                        save_row(failed_writer, row, failed_file)
                        row_already_written = True
                        continue

                    if save_result == "required_value_missing":
                        contents = handle_required_value_missing()

                        if contents:
                            print(f"Job {job}: Required field missing and first field not empty. Marked as failed.")
                            save_row(failed_writer, row, failed_file)
                            row_already_written = True
                            close_without_saving()
                            continue
                        else:
                            print(f"Job {job}: Empty line detected. Deleting empty line and closing.")
                            delete_empty_line_and_close()
                            save_row(failed_writer, row, failed_file)
                            row_already_written = True
                            continue

                    ########################################
                    # SUCCESS
                    ########################################
                    print(f"Job {job}: Success.")
                    save_row(succeeded_writer, row, succeeded_file)
                    row_already_written = True

                    activate_window(SOFTSHIP_WINDOW)
                    click_in_window(SOFTSHIP_WINDOW, CLOSE_BTN)

                except StopExecution:
                    print(f"\nExecution interrupted during row {row_number}: Job={job}")

                    # Current row counts as failed if not already written
                    if not row_already_written:
                        save_row(failed_writer, row, failed_file)
                        row_already_written = True

                    # Append all remaining unprocessed rows to failed
                    for remaining_row in all_rows[idx + 1:]:
                        save_row(failed_writer, remaining_row, failed_file)

                    print("Remaining unprocessed rows appended to failed file.")
                    break

                except Exception as e:
                    print(f"Job {job}: Unexpected error: {e}")

                    if not row_already_written:
                        save_row(failed_writer, row, failed_file)
                        row_already_written = True

            failed_file.flush()
            succeeded_file.flush()

        # if STOP_REQUESTED:
        #     messagebox.showinfo(
        #         "Stopped",
        #         f"Execution was stopped by pressing ESC.\n\n"
        #         f"Failed rows saved to:\n{failed_path}\n\n"
        #         f"Succeeded rows saved to:\n{succeeded_path}"
        #     )
        # else:
        #     messagebox.showinfo(
        #         "Done",
        #         f"Processing complete.\n\n"
        #         f"Failed rows saved to:\n{failed_path}\n\n"
        #         f"Succeeded rows saved to:\n{succeeded_path}"
        #     )

        print("\nDone.")
        print(f"Failed rows saved to: {failed_path}")
        print(f"Succeeded rows saved to: {succeeded_path}")

    except StopExecution:
        messagebox.showinfo("Stopped", "Execution was stopped before processing began.")
        print("Execution was stopped before processing began.")

    except ValueError as e:
        messagebox.showerror("CSV Validation Error", str(e))
        print(f"CSV Validation Error: {e}")

    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))
        print(f"Unexpected Error: {e}")

    finally:
        keyboard.unhook_all_hotkeys()


########################################
# ENTRY POINT
########################################

if __name__ == "__main__":
    main()