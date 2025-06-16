import tkinter as tk
from tkinter import filedialog
import pyautogui
import time
import os
from datetime import datetime

class CaptureTool:
    def __init__(self, root):
        self.root = root
        self.root.attributes("-topmost", True)
        self.root.overrideredirect(True)
        self.root.geometry("300x200+100+100")  # 初期サイズ
        self.root.configure(bg="red")
        self.root.wm_attributes("-alpha", 0.4)

        self.frame = tk.Frame(self.root, bg="black", bd=2)
        self.frame.pack(fill="both", expand=True)

        self.add_resize_grips()
        self.enable_drag()

    def add_resize_grips(self):
        size = 10
        positions = {"nw": "top_left_corner", "ne": "top_right_corner", "sw": "bottom_left_corner", "se": "bottom_right_corner"}
        for pos, cursor in positions.items():
            grip = tk.Frame(self.frame, bg="gray", width=size, height=size, cursor=cursor)
            grip.place(relx=0 if "w" in pos else 1, rely=0 if "n" in pos else 1, anchor=pos)
            grip.bind("<B1-Motion>", self.resize_window)

    def resize_window(self, event):
        self.root.geometry(f"{event.x_root}x{event.y_root}")

    def enable_drag(self):
        self.frame.bind("<ButtonPress-1>", self.start_move)
        self.frame.bind("<B1-Motion>", self.on_move)

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def on_move(self, event):
        x_offset = event.x_root - self.x
        y_offset = event.y_root - self.y
        self.root.geometry(f"+{x_offset}+{y_offset}")

class ControlPanel:
    def __init__(self, root):
        self.root = root
        self.root.title("スクリーンショットツール")
        self.root.geometry("400x200")

        # 撮影領域ウィンドウを作成（先に作ることで座標取得が可能）
        self.capture_window = tk.Toplevel()
        self.capture_tool = CaptureTool(self.capture_window)

        # 撮影領域の座標を取得
        self.root.update_idletasks()
        capture_x, capture_y = self.capture_window.winfo_x(), self.capture_window.winfo_y()
        capture_width, capture_height = 300, 200  # 初期サイズ (上で設定した値)

        # 画面サイズ取得
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 操作パネルの位置計算（右側 or 下側に配置）
        panel_x = capture_x + capture_width + 20  # 右側に配置
        panel_y = capture_y

        if panel_x + 400 > screen_width:  # 右側がはみ出すなら下側へ
            panel_x = capture_x
            panel_y = capture_y + capture_height + 20

        if panel_y + 200 > screen_height:  # 下側もはみ出すなら左上へ
            panel_x, panel_y = 50, 50

        self.root.geometry(f"400x200+{panel_x}+{panel_y}")  # 操作パネルの配置

        # 撮影回数入力
        tk.Label(root, text="撮影回数:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.capture_count = tk.Entry(root)
        self.capture_count.grid(row=0, column=1, padx=5, pady=5)
        self.capture_count.insert(0, "10")

        # 撮影間隔入力
        tk.Label(root, text="撮影間隔(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.interval = tk.Entry(root)
        self.interval.grid(row=1, column=1, padx=5, pady=5)
        self.interval.insert(0, "2")

        # 保存フォルダ選択
        tk.Label(root, text="保存フォルダ:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.folder_path = tk.Entry(root, width=30)
        self.folder_path.grid(row=2, column=1, padx=5, pady=5)
        self.folder_path.insert(0, "C:/Screenshots")

        tk.Button(root, text="参照", command=self.select_folder).grid(row=2, column=2, padx=5, pady=5)

        # 撮影開始・終了ボタン
        tk.Button(root, text="撮影開始", command=self.start_capture, bg="green", fg="white").grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")
        tk.Button(root, text="終了", command=self.root.quit, bg="red", fg="white").grid(row=3, column=2, pady=10, sticky="ew")

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.delete(0, tk.END)
            self.folder_path.insert(0, folder_selected)

    def start_capture(self):
        count = int(self.capture_count.get())
        interval = float(self.interval.get())
        folder = self.folder_path.get()

        if not os.path.exists(folder):
            os.makedirs(folder)

        x, y, width, height = (
            self.capture_window.winfo_x(),
            self.capture_window.winfo_y(),
            self.capture_window.winfo_width(),
            self.capture_window.winfo_height(),
        )

        for _ in range(count):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(folder, f"screenshot_{timestamp}.png")

            self.capture_window.withdraw()
            self.root.withdraw()

            time.sleep(0.1)

            screenshot = pyautogui.screenshot(region=(x, y, width, height))
            screenshot.save(filename)

            self.capture_window.deiconify()
            self.root.deiconify()

            print(f"保存: {filename}")
            time.sleep(interval)

        print("スクリーンショット完了！")

if __name__ == "__main__":
    root = tk.Tk()
    app = ControlPanel(root)
    root.mainloop()
