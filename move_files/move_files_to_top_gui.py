# 指定フォルダ以下を検索し、サブフォルダ内にあるファイルを
# 親フォルダへ集約する

import os
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

def move_files_to_parent(parent_folder):
    parent_folder = os.path.abspath(parent_folder)

    if not os.path.isdir(parent_folder):
        messagebox.showerror("エラー", f"指定したフォルダが存在しません:\n{parent_folder}")
        return

    moved_count = 0
    failed_count = 0
    deleted_folder_count = 0

    for root, dirs, files in os.walk(parent_folder, topdown=False):
        for file in files:
            src_path = os.path.join(root, file)
            dst_path = os.path.join(parent_folder, file)

            base, ext = os.path.splitext(file)
            counter = 1
            while os.path.exists(dst_path):
                dst_path = os.path.join(parent_folder, f"{base}_{counter}{ext}")
                counter += 1

            try:
                shutil.move(src_path, dst_path)
                moved_count += 1
            except Exception as e:
                print(f"Failed to move {src_path}: {e}")
                failed_count += 1

        # フォルダが空なら削除（ファイルの有無に関係なく）
        if root != parent_folder:
            try:
                os.rmdir(root)
                deleted_folder_count += 1
            except OSError:
                # フォルダが空でない場合や削除できない場合はスキップ
                pass

    message = (
        f"処理完了！\n\n"
        f"移動したファイル数: {moved_count}\n"
        f"失敗したファイル数: {failed_count}\n"
        f"削除した空フォルダ数: {deleted_folder_count}"
    )
    messagebox.showinfo("完了", message)

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_folder_path.delete(0, tk.END)
        entry_folder_path.insert(0, folder_selected)

def start_process():
    parent_folder_path = entry_folder_path.get()
    if not parent_folder_path:
        messagebox.showwarning("警告", "フォルダを選択してください。")
        return
    move_files_to_parent(parent_folder_path)

# --- GUI ---
root = tk.Tk()
root.title("ファイル移動プログラム")
root.geometry("500x150")

frame = tk.Frame(root)
frame.pack(pady=20)

entry_folder_path = tk.Entry(frame, width=50)
entry_folder_path.pack(side=tk.LEFT, padx=5)

btn_browse = tk.Button(frame, text="参照", command=select_folder)
btn_browse.pack(side=tk.LEFT)

btn_start = tk.Button(root, text="実行", command=start_process, width=20)
btn_start.pack(pady=10)

root.mainloop()
