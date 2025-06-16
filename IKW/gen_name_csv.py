import xml.etree.ElementTree as ET
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def xml_to_csv_multiple_tables(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        tables = root.findall('Table')
        if not tables:
            raise ValueError("Tableタグが見つかりません。")

        base_name = os.path.splitext(os.path.basename(xml_path))[0]
        base_dir = os.path.dirname(xml_path)

        for idx, table in enumerate(tables, start=1):
            # ヘッダー取得
            thead = table.find('THead')
            if thead is None:
                raise ValueError(f"{idx}番目のTableにTHeadタグが見つかりません。")
            header_row = thead.find('Row')
            headers = [cell.text if cell.text else "" for cell in header_row.findall('Cell')]

            # データ取得（TBody ← 大文字Bに対応）
            tbody = table.find('TBody')
            if tbody is None:
                raise ValueError(f"{idx}番目のTableにTBodyタグが見つかりません。")
            data = []
            for row in tbody.findall('Row'):
                cells = [cell.text if cell.text else "" for cell in row.findall('Cell')]
                data.append(cells)

            # 出力ファイル名：元ファイル名 + _table1.csv など
            csv_filename = f"{base_name}_table{idx}.csv"
            csv_path = os.path.join(base_dir, csv_filename)

            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(data)

        messagebox.showinfo("成功", f"{len(tables)}個のTableをCSVに出力しました。")
    except Exception as e:
        messagebox.showerror("エラー", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("XML files", "*.xml")],
        title="XMLファイルを選択"
    )
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)

def convert_file():
    xml_path = entry_path.get()
    if not xml_path or not os.path.exists(xml_path):
        messagebox.showerror("エラー", "有効なXMLファイルを選択してください。")
        return
    xml_to_csv_multiple_tables(xml_path)

# GUI構築
root = tk.Tk()
root.title("XML → CSV 変換ツール")
root.geometry("500x150")

frame = tk.Frame(root)
frame.pack(pady=10)

entry_path = tk.Entry(frame, width=50)
entry_path.pack(side=tk.LEFT, padx=5)

btn_browse = tk.Button(frame, text="参照...", command=browse_file)
btn_browse.pack(side=tk.LEFT)

btn_convert = tk.Button(root, text="CSVファイル出力", command=convert_file, height=2)
btn_convert.pack(pady=10)

root.mainloop()
