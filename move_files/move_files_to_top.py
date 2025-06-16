import os
import shutil
import sys

def move_files_to_parent(parent_folder):
    # 親フォルダの絶対パスを取得
    parent_folder = os.path.abspath(parent_folder)

    # 親フォルダが存在するかチェック
    if not os.path.isdir(parent_folder):
        print(f"Error: 指定したフォルダが存在しません: {parent_folder}")
        sys.exit(1)

    for root, dirs, files in os.walk(parent_folder, topdown=False):
        for file in files:
            src_path = os.path.join(root, file)
            dst_path = os.path.join(parent_folder, file)

            # 同名ファイルがある場合、リネームして回避
            base, ext = os.path.splitext(file)
            counter = 1
            while os.path.exists(dst_path):
                dst_path = os.path.join(parent_folder, f"{base}_{counter}{ext}")
                counter += 1

            try:
                shutil.move(src_path, dst_path)
                print(f"Moved: {src_path} -> {dst_path}")
            except Exception as e:
                print(f"Failed to move {src_path}: {e}")

        # ファイル移動後、空になったフォルダを削除
        if root != parent_folder:  # 親フォルダ自身は削除しない
            try:
                os.rmdir(root)
                print(f"Deleted empty folder: {root}")
            except OSError:
                # フォルダが空でない場合はスキップ
                pass

if __name__ == "__main__":
    # ここに親フォルダのパスを指定
    parent_folder_path = r"C:\Users\P0244\Downloads\xmlaaaa"
    move_files_to_parent(parent_folder_path)
