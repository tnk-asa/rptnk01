import pyautogui
import time
import sys
import os
import tkinter as tk
import configparser

## ----- 初期処理：カレントディレクトリ -----
def get_dir_path(relative_path):
    try:
        base_path = sys._MEIPASS
        print("[Base Path (get from sys)]" + base_path)
    except Exception:
        base_path = os.path.dirname(__file__)
        print("[Base Path (get from sys)]" + base_path)
    return os.path.join(base_path, relative_path)
# -----
current_file = get_dir_path(sys.argv[0])
current_dir = os.path.dirname(os.path.abspath(current_file))
os.chdir(current_dir)

# 設定ファイル読み込み
config = configparser.ConfigParser()
config_path = "config.ini"

if not os.path.exists(config_path):
    # 初回実行時にデフォルトの `config.ini` を作成
    config["Settings"] = {
        "save_path": "C:/Screenshots",
        "max_shots": "10",
        "interval": "2"
    }
    with open(config_path, "w") as configfile:
        config.write(configfile)

config.read(config_path)

# 設定値の取得
save_path = config.get("Settings", "save_path", fallback="C:/Screenshots")
max_shots = config.getint("Settings", "max_shots", fallback=10)
interval = config.getint("Settings", "interval", fallback=2)

# 初期サイズと位置
x, y, width, height = 100, 100, 800, 600

# Windows 用カーソル設定
CURSOR_MAP = {
    "nw": "size_nw_se",
    "ne": "size_ne_sw",
    "sw": "size_nw_se",
    "se": "size_ne_sw",
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

        self.frame = tk.Frame(self.root, bg="black", bd=2)
        self.frame.pack(fill="both", expand=True)

        self.start_x = self.start_y = None
        self.resizing = False
        self.moving = False

        self.frame.bind("<ButtonPress-1>", self.start_move)
        self.frame.bind("<B1-Motion>", self.on_move)
        self.frame.bind("<ButtonRelease-1>", self.stop_move)

        self.add_resize_grips()

        # **ボタンエリア**
        button_frame = tk.Frame(self.frame, bg="black")
        button_frame.pack(side="top", fill="x")

        # **確定ボタン**
        self.confirm_button = tk.Button(button_frame, text="撮影開始", command=self.start_screenshot, bg="green", fg="white")
        self.confirm_button.pack(side="left", expand=True, fill="x")

        # **終了ボタン**
        self.exit_button = tk.Button(button_frame, text="終了", command=self.exit_program, bg="red", fg="white")
        self.exit_button.pack(side="right", expand=True, fill="x")

    def exit_program(self):
        """ウィンドウを閉じて終了"""
        self.root.destroy()

    def add_resize_grips(self):
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
            cursor_type = CURSOR_MAP.get(pos, "arrow")  
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
