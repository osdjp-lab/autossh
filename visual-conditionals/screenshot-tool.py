import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageGrab


class AreaScreenshotTool:
    def __init__(self, root):
        self.root = root
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.25)
        self.root.configure(bg="gray")

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="gray", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        self.root.bind("<Escape>", self.cancel)

        self.info_label = tk.Label(
            self.root,
            text="Click and drag to select an area. Press Esc to cancel.",
            bg="black",
            fg="white",
            font=("Arial", 12)
        )
        self.info_label.place(x=20, y=20)

    def on_mouse_down(self, event):
        self.start_x = event.x
        self.start_y = event.y

        if self.rect:
            self.canvas.delete(self.rect)

        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="red", width=2
        )

    def on_mouse_drag(self, event):
        if self.rect:
            self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_mouse_up(self, event):
        end_x = event.x
        end_y = event.y

        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)

        width = x2 - x1
        height = y2 - y1

        if width < 2 or height < 2:
            messagebox.showwarning("Too small", "Selected area is too small.")
            return

        self.root.withdraw()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            title="Save screenshot as..."
        )

        if not file_path:
            self.root.destroy()
            return

        screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        screenshot.save(file_path)

        self.root.destroy()
        messagebox.showinfo("Saved", f"Screenshot saved to:\n{file_path}")

    def cancel(self, event=None):
        self.root.destroy()


def main():
    root = tk.Tk()
    app = AreaScreenshotTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
