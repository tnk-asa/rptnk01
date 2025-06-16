import xml.etree.ElementTree as ET
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

def convert_class_structure(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        rows = []
        all_kinds = set(['kind', 'code', 'SuperClassCode'])  # 初期見出しセット

        for class_elem in root.findall('Class'):
            row = {}

            # Class の kind 属性
            kind = class_elem.get('kind')
            if kind:
                row['kind'] = kind

            # Class の code 属性
            class_code = class_elem.get('code')
            if class_code:
                row['code'] = class_code

            # SuperClass の code 属性
            superclass = class_elem.find('SuperClass')
            if superclass is not None:
                superclass_code = superclass.get('code')
                if superclass_code:
                    row['SuperClassCode'] = superclass_code

            # Rubric タグ処理（Label含む全テキストを取得）
            for rubric in class_elem.findall('Rubric'):
                rubric_kind = rubric.get('kind')
                if rubric_kind:
                    text = " ".join(t.strip() for t in rubric.itertext() if t.strip())
                    row[rubric_kind] = text
                    all_kinds.add(rubric_kind)

            rows.append(row)

        # 出力用ヘッダー（kind順で並べ替え、必須列を先頭に固定）
        fixed_headers = ['kind', 'code', 'SuperClassCode']
        dynamic_headers = sorted(all_kinds - set(fixed_headers))
        header = fixed_headers + dynamic_headers

        # 出力パス決定
        base_name = os.path.splitext(os.path.basename(xml_path))[0]
        csv_path = os.path.join(os.path.dirname(xml_path), base_name + "_class_output.csv")

        # CSV書き出し（UTF-8 BOM付き）
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=header)
            writer.writeheader()
            for row in rows:
                writer.writerow({k: row.get(k, '') for k in header})

        messagebox.showinfo("成功", f"CSVファイルを出力しました:\n{csv_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")

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
    convert_class_structure(xml_path)

# GUI 構築
root = tk.Tk()
root.title("Class形式XML → CSV変換ツール")
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
