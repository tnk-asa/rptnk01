import pyautogui
import time
import os
import tkinter as tk

# 設定（スクリーンショット保存先、回数、間隔）
save_path = "C:/Screenshots"
max_shots = 6  # 撮影回数
interval = 3  # 撮影間隔（秒）

# 初期サイズと位置
x, y, width, height = 100, 100, 800, 600

class CaptureTool:
    def __init__(self, root):
        self.root = root
        self.root.attributes("-topmost", True)  # 最前面表示
        self.root.overrideredirect(True)  # ウィンドウ枠を非表示
        self.root.geometry(f"{width}x{height}+{x}+{y}")  # 初期位置
        self.root.configure(bg="red")
        self.root.wm_attributes("-alpha", 0.3)  # 半透明

        # ドラッグ・リサイズ用変数
        self.start_x = self.start_y = None
        self.dragging = False

        # イベント設定
        self.root.bind("<ButtonPress-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.on_drag)
        self.root.bind("<ButtonRelease-1>", self.stop_drag)
        self.root.bind("<Escape>", lambda e: self.root.destroy())  # ESCキーで終了

        # 確定ボタン（エリア上部に配置）
        self.confirm_button = tk.Button(self.root, text="確定", command=self.start_screenshot, bg="black", fg="white")
        self.confirm_button.pack(side="top", fill="x")

    def start_drag(self, event):
        self.start_x, self.start_y = event.x_root, event.y_root
        self.dragging = True

    def on_drag(self, event):
        if self.dragging:
            dx, dy = event.x_root - self.start_x, event.y_root - self.start_y
            new_x = self.root.winfo_x() + dx
            new_y = self.root.winfo_y() + dy
            self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}+{new_x}+{new_y}")
            self.start_x, self.start_y = event.x_root, event.y_root

    def stop_drag(self, event):
        self.dragging = False

    def start_screenshot(self):
        # 現在のウィンドウ位置を取得
        x, y = self.root.winfo_x(), self.root.winfo_y()
        width, height = self.root.winfo_width(), self.root.winfo_height()

        self.root.destroy()  # 確定後に枠を消す
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
