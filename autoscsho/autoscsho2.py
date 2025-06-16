import pyautogui
import time
import os
import tkinter as tk

# 設定
save_path = "C:/Screenshots"
max_shots = 10  
interval = 1  

# 初期サイズと位置
x, y, width, height = 100, 100, 800, 600

# Windows 用のカーソル設定
CURSOR_MAP = {
    "nw": "size_nw_se",
    "ne": "size_ne_sw",
    "sw": "size_nw_se",  # 修正: sw も size_nw_se に統一
    "se": "size_ne_sw",  # 修正: se も size_ne_sw に統一
    "n": "size_ns",
    "s": "size_ns",
    "w": "size_we",
    "e": "size_we",
}

class CaptureTool:
    def __init__(self, root):
        self.root = root
        self.root.attributes("-topmost", True)  
        self.root.overrideredirect(True)  
        self.root.geometry(f"{width}x{height}+{x}+{y}")  
        self.root.configure(bg="red")
        self.root.wm_attributes("-alpha", 0.4)  

        # フレーム（枠線）を表示
        self.frame = tk.Frame(self.root, bg="black", bd=2)
        self.frame.pack(fill="both", expand=True)

        # ドラッグ・リサイズ用変数
        self.start_x = self.start_y = None
        self.resizing = False
        self.moving = False

        # イベント設定
        self.frame.bind("<ButtonPress-1>", self.start_move)
        self.frame.bind("<B1-Motion>", self.on_move)
        self.frame.bind("<ButtonRelease-1>", self.stop_move)

        # リサイズ用のグリップを追加
        self.add_resize_grips()

        # 確定ボタン（エリア上部に配置）
        self.confirm_button = tk.Button(self.frame, text="確定", command=self.start_screenshot, bg="black", fg="white")
        self.confirm_button.pack(side="top", fill="x")

    def add_resize_grips(self):
        """四隅と辺にリサイズ用の小さい枠を追加"""
        size = 10  
        self.grips = []

        positions = [
            ("nw", "top", "left"),
            ("ne", "top", "right"),
            ("sw", "bottom", "left"),
            ("se", "bottom", "right"),
            ("n", "top", "center"),
            ("s", "bottom", "center"),
            ("w", "center", "left"),
            ("e", "center", "right")
        ]

        for pos, side_v, side_h in positions:
            cursor_type = CURSOR_MAP.get(pos, "arrow")  # 無効なカーソルは "arrow" にする
            grip = tk.Frame(self.frame, bg="gray", width=size, height=size, cursor=cursor_type)
            grip.place(relx={"left": 0, "center": 0.5, "right": 1}[side_h], 
                       rely={"top": 0, "center": 0.5, "bottom": 1}[side_v], 
                       anchor=pos)

            grip.bind("<ButtonPress-1>", lambda e, p=pos: self.start_resize(e, p))
            grip.bind("<B1-Motion>", lambda e, p=pos: self.on_resize(e, p))
            grip.bind("<ButtonRelease-1>", self.stop_resize)
            self.grips.append(grip)

    def start_move(self, event):
        self.start_x, self.start_y = event.x_root, event.y_root
        self.moving = True

    def on_move(self, event):
        if self.moving:
            dx, dy = event.x_root - self.start_x, event.y_root - self.start_y
            new_x = self.root.winfo_x() + dx
            new_y = self.root.winfo_y() + dy
            self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}+{new_x}+{new_y}")
            self.start_x, self.start_y = event.x_root, event.y_root

    def stop_move(self, event):
        self.moving = False

    def start_resize(self, event, pos):
        self.start_x, self.start_y = event.x_root, event.y_root
        self.resizing = True
        self.resize_direction = pos

    def on_resize(self, event, pos):
        if self.resizing:
            dx, dy = event.x_root - self.start_x, event.y_root - self.start_y
            x, y, w, h = self.root.winfo_x(), self.root.winfo_y(), self.root.winfo_width(), self.root.winfo_height()

            if "w" in self.resize_direction:
                x += dx
                w -= dx
            if "e" in self.resize_direction:
                w += dx
            if "n" in self.resize_direction:
                y += dy
                h -= dy
            if "s" in self.resize_direction:
                h += dy

            if w > 50 and h > 50:  
                self.root.geometry(f"{w}x{h}+{x}+{y}")

            self.start_x, self.start_y = event.x_root, event.y_root

    def stop_resize(self, event):
        self.resizing = False

    def start_screenshot(self):
        x, y = self.root.winfo_x(), self.root.winfo_y()
        width, height = self.root.winfo_width(), self.root.winfo_height()

        self.root.destroy()  
        self.capture_screenshots(x, y, width, height)

    def capture_screenshots(self, x, y, width, height):
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        for i in range(1, max_shots + 1):
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(save_path, f"Screenshot_{timestamp}.png")

            screenshot = pyautogui.screenshot(region=(x, y, width, height))
            screenshot.save(filename)

            print(f"[{i}/{max_shots}] Saved: {filename} (X={x}, Y={y}, W={width}, H={height})")

            if i < max_shots:
                time.sleep(interval)

        print("Screenshot capture completed.")

# GUI起動
root = tk.Tk()
app = CaptureTool(root)
root.mainloop()
