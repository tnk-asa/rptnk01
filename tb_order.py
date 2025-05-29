import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES  # ドラッグ＆ドロップ用
from datetime import datetime, timedelta
import os
from pathlib import Path
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter

# 版素コードリスト
hanso_code_list = ["", 
    "1-A1","1-A2","1-A3","1-A4","1-A5","1-A6",
    "1-B1","1-B2","1-B3","1-B4","1-B5","1-B6",
    "1-C1","1-C2","1-C3","1-C4","1-C5","1-C6",
    "2-A1","2-A2","2-A3","2-A4","2-A5","2-A6",
    "2-B1","2-B2","2-B3","2-B4","2-B5","2-B6",
    "2-C1","2-C2","2-C3","2-C4","2-C5","2-C6",
    "2-D1","2-D2","2-D3","2-D4","2-D5","2-D6",
    "2-E1","2-E2","2-E3","2-E4","2-E5","2-E6",
    # "KRS-A1","NTD-A1","OLS-A1",
    # "SP1-B1","SP1-B2","SP1-B3","SP1-B4","SP1-B5","SP1-B6",
    # "SP2-A1","SP2-A2","SP2-A3","SP2-A4","SP2-A5","SP2-A6",
    # "SP2-B1","SP2-B2","SP2-B3","SP2-B4","SP2-B5","SP2-B6",
    # "SP2-C1","SP2-C2","SP2-C3","SP2-C4","SP2-C5","SP2-C6",
    # "SP2-D1","SP2-D2","SP2-D3","SP2-D4","SP2-D5","SP2-D6",
    # "SP2-E1","SP2-E2","SP2-E3","SP2-E4","SP2-E5","SP2-E6",
    "SS1-A1","SS1-A2","SS1-A3","SS1-A4","SS1-A5","SS1-A6",
    "SS2-A1","SS2-A2","SS2-A3","SS2-A4","SS2-A5","SS2-A6",
    "SS2-B1","SS2-B2","SS2-B3","SS2-B4","SS2-B5","SS2-B6",
    "SS2-C1","SS2-C2","SS2-C3","SS2-C4","SS2-C5","SS2-C6",
    "SS2-D1","SS2-D2","SS2-D3","SS2-D4","SS2-D5","SS2-D6",
    "SS2-E1","SS2-E2","SS2-E3","SS2-E4","SS2-E5","SS2-E6",
    "SW-A1","SW-A2","SW-A3","SW-A4","SW-A5","SW-A6",
    "SW-B1","SW-B2","SW-B3","SW-B4","SW-B5","SW-B6",
    "SW-C1","SW-C2","SW-C3","SW-C4","SW-C5","SW-C6","SW-S1",
    "SX1-C1","SX1-C2","SX1-C3","SX1-C4","SX1-C5","SX1-C6",
    "SX2-A1","SX2-A2","SX2-A3","SX2-A4","SX2-A5","SX2-A6",
    "SX2-B1","SX2-B2","SX2-B3","SX2-B4","SX2-B5","SX2-B6",
    "SX2-C1","SX2-C2","SX2-C3","SX2-C4","SX2-C5","SX2-C6",
    "SX2-D1","SX2-D2","SX2-D3","SX2-D4","SX2-D5","SX2-D6",
    "SX2-E1","SX2-E2","SX2-E3","SX2-E4","SX2-E5","SX2-E6"    
]

# PDFファイルにテキストを追加
def add_text_to_pdf(input_pdf_path, output_pdf_path, out_text, x_coordinate, y_coordinate, item_class, RGB_list):
    pdf_document = fitz.open(input_pdf_path)
    # 追加するテキストの内容とスタイル
    text = out_text # 出力するテキスト
    text_color = (RGB_list[0], RGB_list[1], RGB_list[2])  # 青 (RGB, 0-1 の範囲)
    background_color = (1, 1, 1)  # 白 (RGB, 0-1 の範囲)
    border_color = (RGB_list[0], RGB_list[1], RGB_list[2])  # 赤 (1,1,0 RGB, 0-1 の範囲)
    font_size = 18
    fontname = "japan"
    # PDFの1ページ目にテキストを追加
    page = pdf_document[0]
    # 項目ごとの位置調整
    if item_class == "1" : # 受付情報
        relative_x = 380
        relative_y = 270
    elif item_class == "2" : # 版素コード
        relative_x = 260
        relative_y = 160
    elif item_class == "3" : # その他連絡事項
        relative_x = 600
        relative_y = 50
    elif item_class == "4" : # 振分コード
        relative_x = 80
        relative_y = 50
    else :
        relative_x = 320
        relative_y = 28
        border_color = (1, 1, 1)
        
    # テキストの位置とエリアサイズを設定
    text_position = fitz.Point(x_coordinate, y_coordinate)
    text_rect = fitz.Rect(
        text_position.x, text_position.y,
        text_position.x + relative_x, text_position.y + relative_y
    ) 
    # テキストエリアの背景
    page.draw_rect(text_rect, color=background_color, fill=background_color)
    # テキストエリアの境界線
    page.draw_rect(text_rect, color=border_color, width=1)
    # テキストをエリア内に描画
    page.insert_textbox(
        text_rect,
        text,
        fontsize=font_size,
        color=text_color,
        fontname=fontname,
        rotate=0, 
        align=0  # 1:中央揃え
    )

    # 変更を保存
    pdf_document.save(output_pdf_path)
    pdf_document.close()

# PDFの先頭ページだけ残す
def trim_pdf_to_first_page(pdf_path: str, output_path: str = None):
    # PDFのリーダーとライターを初期化
    reader = PdfReader(pdf_path)
    # PDFが1ページだけの場合は処理しない
    if len(reader.pages) == 1:
        return
    # 1ページ目を新しいPDFに追加
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    # 出力先を指定されなければ元ファイルを上書き
    output_path = output_path if output_path else pdf_path
    # ファイルを書き込み
    with open(output_path, "wb") as out_file:
        writer.write(out_file)

def duplicate_pdf_page(pdf_path, output_path): # 1ページ目ｘ2を別ファイルで保存
    original_pdf = fitz.open(pdf_path)
    conv_pdf = fitz.open()
    # 指定したページを複製
    for _ in range(2):
        conv_pdf.insert_pdf(original_pdf, from_page=0, to_page=0)
    # 新しいPDFとして保存
    conv_pdf.save(output_path)
    original_pdf.close()
    conv_pdf.close()

# 入力内容を表示する関数
def display_inputs(production, dividecode, date, time_selection, category_var, 
                   data_creation, font_update, hanso1, hanso2, hanso3, hanso4, 
                   check_tombo, check_mentuke, check_aidata, check_txtaidata, announce_text, file_path):
    result = f"PDFに情報を追加しました。:\n\n"
    result += f"制作先: {production}\n"
    result += f"制作先: {dividecode}\n"
    result += f"日付: {date}\n"
    result += f"時間帯: {time_selection}\n"
    result += f"区分: {category_var}\n"
    result += f"支給データ通り: {data_creation}\n"
    result += f"フォントバージョンアップ: {font_update}\n"
    result += f"版素コード: {hanso1} {hanso2} {hanso3} {hanso4}\n"    
    if check_tombo or check_mentuke or check_aidata:
        result += "同時作成:\n"
        if check_tombo:
            result += "  - トンボ無PDF\n"
        if check_mentuke:
            result += "  - 面付PDF\n"
        if check_aidata:
            result += "  - AIデータ\n"

    result += f"連絡事項: {announce_text}\n"
    result += f"ファイルパス: {file_path}\n"

    # メッセージボックスで表示
    messagebox.showinfo("入力結果", result)

def tenbun_order_gui():
    def increment_date():  # 日付＋1日
        current_date = datetime.strptime(date_var.get(), '%Y-%m-%d')
        new_date = current_date + timedelta(days=1)
        date_var.set(new_date.strftime('%Y-%m-%d'))
    def decrement_date():  # 日付－1日
        current_date = datetime.strptime(date_var.get(), '%Y-%m-%d')
        new_date = current_date - timedelta(days=1)
        date_var.set(new_date.strftime('%Y-%m-%d'))
    def set_today(): # 日付セット
        today_date = datetime.now().strftime('%Y-%m-%d')
        date_var.set(today_date)
    def browse_file(): # ファイル選択ダイアログ
        file_path = filedialog.askopenfilename(title="ファイルを選択")
        file_var.set(file_path)
    def drop_file(event): # ファイル追加Ｄ＆Ｄ対応
        file_var.set(event.data.strip('{}'))
        
    def on_execute(): # 実行ボタンが押された時の処理
        # 入力チェック
        if  file_var.get() == "" or file_var.get() == None:
            messagebox.showerror("ファイル選択エラー","ファイルが選択されていません")
            return
        file_ext_pair = os.path.splitext(file_var.get()) # 拡張子チェック
        if file_ext_pair[1] != '.pdf' :
            messagebox.showerror("ファイル種類エラー","PDFではないファイルが選択されています")
            return
        if  hanso1_var.get() == "" and hanso2_var.get() == "" and hanso3_var.get() == "" and hanso4_var.get() == "" :
            messagebox.showerror("版素コードエラー","版素コードが１つも入力されていません")
            return
        if dividecode_var.get() == "65" or dividecode_var.get() == "67":
            if production_var.get() != "添文内作" :
                messagebox.showerror("振分コードエラー","制作先と振分コードの組み合わせが間違っています")
                return
        elif production_var.get() == "添文内作" :
            messagebox.showerror("振分コードエラー","制作先と振分コードの組み合わせが間違っています")
            return
                    
        # 入力ファイルのディレクトリへ移動
        pdf_path = file_var.get()
        os.chdir(os.path.dirname(os.path.abspath(pdf_path)))
        
        # 出力ファイル名生成
        month_day = datetime.now().strftime("%m%d") #MMDD
        input_file = Path(pdf_path)
        file_name_org = input_file.stem # 元ファイル名
        file_ext_org  = input_file.suffix # 拡張子
        #file_dir_org  = input_file.parent # ディレクトリ
#        pdf_path_split = os.path.split(pdf_path)
#        output_pdf = production_var.get() + "_" + pdf_path_split[1] + "_" + month_day
        output_pdf = production_var.get() + "_" + file_name_org + "_" + month_day + file_ext_org
        #print(output_pdf)
        if os.path.exists(output_pdf): # 同じファイル名が存在する場合は中断
            messagebox.showinfo("Message","既に受付情報追加済みのPDFが存在します")
            return
        
        deadline_date = str(date_var.get())
        deadline_time = ""
        #print(deadline_date)
        if time_var.get() != "指定なし" and time_var.get() != "" and time_var.get() != None :
            deadline_time = time_var.get()

        # 出力用テキスト１　編集
        output_text = f"{production_var.get()}"
        if production_var.get() != "添文内作" :
            output_text += f"　御中\n{deadline_date} {deadline_time} \nまでにデータを送ってください\n"

        if derivation_tombo.get() == 1 or derivation_mentuke.get() == 1 or derivation_aidata.get() == 1  or derivation_txtaidata.get() == 1 :
            output_text += "\n・同時作成\n"
            if derivation_tombo.get() == 1 :
                output_text += "--トンボ無PDF\n"
            if derivation_mentuke.get() == 1 :
                output_text += "--面付PDF\n"
            if derivation_aidata.get() == 1 or derivation_txtaidata.get() == 1 :
                output_text += "--AIデータ："
                if derivation_aidata.get() == 1 :
                    output_text += "アウトライン　"
                if derivation_txtaidata.get() == 1 :
                    output_text += "テキスト"
                output_text += "\n"
 
        if data_var.get() != 0 and data_var.get() != "0":
            output_text += "・支給データ通り作成　"
            output_text += f"{data_var.get()}\n" 

        if font_var.get() != 0 and font_var.get() != "0":
            output_text += "・フォントバージョンアップ　"
            output_text += f"{font_var.get()}\n" 

        #output_text += f"制作先: {production_var.get()}\n"        
        # 出力用テキスト　版素コード　編集
        # output_text2 = "版素コード\n"
        output_text2 = f"添文区分：{category_var.get()}\n\n"
        if hanso1_var.get() != "":
            output_text2 += f"{hanso1_var.get()}\n"
        if hanso2_var.get() != "":
            output_text2 += f"{hanso2_var.get()}\n"
        if hanso3_var.get() != "":
            output_text2 += f"{hanso3_var.get()}\n"
        if hanso4_var.get() != "":
            output_text2 += f"\n{hanso4_var.get()}\n"

        # 連絡事項            
        output_text3 = announce_text.get("1.0", "end").strip()            

        # 出力用テキスト　振分コード　編集
        if dividecode_var.get() == "65" :
            output_text4 = "内作一面\n65"
        if dividecode_var.get() == "67" :
            output_text4 = "内作面付\n67"
        if dividecode_var.get() == "60" :
            output_text4 = "外注一面\n60"
        if dividecode_var.get() == "69" :
            output_text4 = "外注面付\n69"
        
        #print(output_text)
        #print(output_text2)
        #print(output_text3)
        #print(output_text4)
        # PDF編集
        duplicate_pdf_page(pdf_path, "tb_order_tmp00.pdf") # 1ページ目のみにする
        # 受付情報
        x_coordinate = 30
        y_coordinate = 360
        RGB_list = [0, 0, 1]
        item_class = "1"
        add_text_to_pdf("tb_order_tmp00.pdf", "tb_order_tmp01.pdf", output_text, x_coordinate, y_coordinate, item_class, RGB_list)
        # 版素コード
        x_coordinate = 30
        y_coordinate = 650
        RGB_list = [1, 0, 0]
        item_class = "2"
        add_text_to_pdf("tb_order_tmp01.pdf", "tb_order_tmp02.pdf", output_text2, x_coordinate, y_coordinate, item_class, RGB_list)
        # 連絡事項
        x_coordinate = 30
        y_coordinate = 920
        RGB_list = [0, 0, 1]
        item_class = "3"
        add_text_to_pdf("tb_order_tmp02.pdf", "tb_order_tmp03.pdf", output_text3, x_coordinate, y_coordinate, item_class, RGB_list)
        # 振分コード
        x_coordinate = 30
        y_coordinate = 100
        RGB_list = [1, 0, 1]
        item_class = "4"
        add_text_to_pdf("tb_order_tmp03.pdf", "tb_order_tmp04.pdf", output_text4, x_coordinate, y_coordinate, item_class, RGB_list)
        # 差出人
        sender_text = "朝日印刷プリプレス部　添文課"
        x_coordinate = 400
        y_coordinate = 30
        RGB_list = [0, 0, 0]
        item_class = "5"
        add_text_to_pdf("tb_order_tmp04.pdf", "tb_order_tmp05.pdf", sender_text, x_coordinate, y_coordinate, item_class, RGB_list)
        # 固定メッセージ
        msg_text = "※2ページ目にオリジナルがあります"
        x_coordinate = 380
        y_coordinate = 980
        RGB_list = [1, 0, 0]
        item_class = "5"
        add_text_to_pdf("tb_order_tmp05.pdf", output_pdf, msg_text, x_coordinate, y_coordinate, item_class, RGB_list)
        
        tmp_files = [f for f in os.listdir() if "tb_order_tmp" in f and os.path.isfile(f)]
        for tmp_file in tmp_files : # 一時ファイル削除
        #for tmp_file in ["tb_order_tmp00.pdf", "tb_order_tmp01.pdf", "tb_order_tmp02.pdf", "tb_order_tmp03.pdf", "tb_order_tmp04.pdf", "tb_order_tmp05.pdf"] : # 一時ファイル削除
            if os.path.exists(tmp_file):
               os.remove(tmp_file)
            else:
               pass
        
        # メッセージボックスで表示
        display_inputs(
            production_var.get(),
            dividecode_var.get(),
            date_var.get(),
            time_var.get(),
            category_var.get(),
            data_var.get(),
            font_var.get(),
            hanso1_var.get(),
            hanso2_var.get(),
            hanso3_var.get(),
            hanso4_var.get(),            
            derivation_tombo.get(),
            derivation_mentuke.get(),
            derivation_aidata.get(),
            derivation_txtaidata.get(),
            announce_text.get("1.0", "end").strip(),
            file_var.get()
        )

        file_var.set("") # ファイルパスクリア
        announce_text.delete("1.0", "end") # 連絡事項をクリア

    #####################################################################################
    # ウィンドウを作成
    root_window = TkinterDnD.Tk()
    root_window.title("添文課　PDFに受付情報を追加")
    root_window.geometry("500x750")

    # ファイル選択のフレーム
    file_frame = ttk.LabelFrame(root_window, text="原版仕様書PDFファイル選択")
    file_frame.pack(pady=10, fill="x", padx=10)

    file_var = tk.StringVar()
    file_entry = ttk.Entry(file_frame, textvariable=file_var, width=40)
    file_entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)

    browse_button = ttk.Button(file_frame, text="参照", command=browse_file)
    browse_button.pack(side="right", padx=5, pady=5)

    file_entry.drop_target_register(DND_FILES)
    file_entry.dnd_bind('<<Drop>>', drop_file)

    # 制作先選択
    production_frame = ttk.LabelFrame(root_window, text="制作先を選択")
    production_frame.pack(pady=10, fill="x", padx=10)

    production_var = tk.StringVar(value="添文内作")
    ttk.Radiobutton(production_frame, text="添文内作",     variable=production_var, value="添文内作").pack(side="left", padx=10)
    ttk.Radiobutton(production_frame, text="シーオーエム", variable=production_var, value="シーオーエム").pack(side="left", padx=10)
    ttk.Radiobutton(production_frame, text="ニッポー",     variable=production_var, value="ニッポー").pack(side="left", padx=10)

    # 振分コード選択
    dividecode_frame = ttk.LabelFrame(root_window, text="振分コードを選択")
    dividecode_frame.pack(pady=10, fill="x", padx=10)

    dividecode_var = tk.StringVar(value="65")
    ttk.Radiobutton(dividecode_frame, text="65：内作一面", variable=dividecode_var, value="65").pack(side="left", padx=10)
    ttk.Radiobutton(dividecode_frame, text="67：内作面付", variable=dividecode_var, value="67").pack(side="left", padx=10)
    ttk.Radiobutton(dividecode_frame, text="60：外注一面", variable=dividecode_var, value="60").pack(side="left", padx=10)
    ttk.Radiobutton(dividecode_frame, text="69：外注面付", variable=dividecode_var, value="69").pack(side="left", padx=10)

    # 日付入力欄
    date_frame = ttk.LabelFrame(root_window, text="納期-日付を入力・時間帯を選択")
    date_frame.pack(pady=10, fill="x", padx=10)

    date_var = tk.StringVar()
    date_var.set(datetime.now().strftime('%Y-%m-%d'))

    #ttk.Label(date_frame, text="日付:").pack(side="left", padx=5)
    ttk.Entry(date_frame, textvariable=date_var, width=12).pack(side="left", padx=5)

    ttk.Button(date_frame, text="<<前の日", width=10, command=decrement_date).pack(side="left", padx=5 )
    ttk.Button(date_frame, text="今日",     width=5, command=set_today).pack(side="left", padx=1)
    ttk.Button(date_frame, text="次の日>>", width=10, command=increment_date).pack(side="left", padx=5)

    # 時間帯選択リスト
    #date_frame = ttk.LabelFrame(root, text="納期-時間帯を選択")
    ttk.Label(date_frame, text="　時間帯:").pack(side="left", padx=1)
    date_frame.pack(pady=5, fill="x", padx=5)

    time_var = tk.StringVar(value="指定なし")
    ttk.Label(date_frame).pack(side="left", padx=5)
    ttk.Combobox(date_frame, width=10, textvariable=time_var, 
                 values=["指定なし", "午前","午後","9:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00"]).pack(side="left", padx=10)
    #ttk.Radiobutton(time_frame, text="午前", variable=time_var, value="午前").pack(side="left", padx=10)
    #ttk.Radiobutton(time_frame, text="午後", variable=time_var, value="午後").pack(side="left", padx=10)
    #ttk.Radiobutton(time_frame, text="指定なし", variable=time_var, value="指定なし").pack(side="left", padx=10)

    # 同時作成-チェックボックス
    derivation_frame = ttk.LabelFrame(root_window, text="同時作成")
    derivation_frame.pack(pady=10, fill="x", padx=10)

    derivation_tombo   = tk.BooleanVar(value=False)
    derivation_mentuke = tk.BooleanVar(value=False)
    derivation_aidata  = tk.BooleanVar(value=False)
    derivation_txtaidata = tk.BooleanVar(value=False)
    ttk.Checkbutton(derivation_frame, text="トンボ無PDF", variable=derivation_tombo).pack(side="left", padx=10)
    ttk.Checkbutton(derivation_frame, text="面付PDF", variable=derivation_mentuke).pack(side="left", padx=10)
    ttk.Checkbutton(derivation_frame, text="AIデータ：アウトライン", variable=derivation_aidata).pack(side="left", padx=10)
    ttk.Checkbutton(derivation_frame, text="AIデータ：テキスト", variable=derivation_txtaidata).pack(side="left", padx=10)    

    # 支給データ通り作成ラジオボタン
    data_frame = ttk.LabelFrame(root_window, text="支給データ通り作成")
    data_frame.pack(pady=10, fill="x", padx=10)

    data_var = tk.StringVar(value=0)
    ttk.Radiobutton(data_frame, text="アウトライン", variable=data_var, value="アウトライン").pack(side="left", padx=10)
    ttk.Radiobutton(data_frame, text="テキスト", variable=data_var, value="テキスト").pack(side="left", padx=10)
    ttk.Radiobutton(data_frame, text="なし", variable=data_var, value=0).pack(side="left", padx=10)

    # フォントVerUPラジオボタン
    font_frame = ttk.LabelFrame(root_window, text="フォントバージョンアップ")
    font_frame.pack(pady=10, fill="x", padx=10)

    font_var = tk.StringVar(value=0)
    ttk.Radiobutton(font_frame, text="ＯＫ", variable=font_var, value="ＯＫ").pack(side="left", padx=10)
    ttk.Radiobutton(font_frame, text="不可", variable=font_var, value="不可").pack(side="left", padx=10)
    ttk.Radiobutton(font_frame, text="なし", variable=font_var, value=0).pack(side="left", padx=10)

    # 添文区分ラジオボタン
    category_frame = ttk.LabelFrame(root_window, text="添文区分")
    category_frame.pack(pady=10, fill="x", padx=10)

    category_var = tk.StringVar(value="その他")
    ttk.Radiobutton(category_frame, text="新記載", variable=category_var, value="新記載").pack(side="left", padx=10)
    ttk.Radiobutton(category_frame, text="その他", variable=category_var, value="その他").pack(side="left", padx=10)

    # 版素コード1　リスト選択
    hanso1_frame = ttk.LabelFrame(root_window, text="版素コード-リスト選択")
    hanso1_frame.pack(pady=10, fill="x", padx=10)
    hanso1_var = tk.StringVar()
    ttk.Label(hanso1_frame).pack(side="left", padx=5)
    ttk.Combobox(hanso1_frame, width=10, textvariable=hanso1_var, values=hanso_code_list).pack(side="left", padx=10)
    # 版素コード2　リスト選択
    #hanso1_frame = ttk.LabelFrame(root_window, text="版素コード２-リスト選択")
    hanso1_frame.pack(pady=10, fill="x", padx=10)
    hanso2_var = tk.StringVar()
    ttk.Label(hanso1_frame).pack(side="left", padx=5)
    ttk.Combobox(hanso1_frame, width=10, textvariable=hanso2_var, values=hanso_code_list).pack(side="left", padx=10)
    # 版素コード3　リスト選択
    #hanso1_frame = ttk.LabelFrame(root_window, text="版素コード３-リスト選択")
    hanso1_frame.pack(pady=10, fill="x", padx=10)
    hanso3_var = tk.StringVar()
    ttk.Label(hanso1_frame).pack(side="left", padx=5)
    ttk.Combobox(hanso1_frame, width=10, textvariable=hanso3_var, values=hanso_code_list).pack(side="left", padx=10)    
    # 特殊版素コード　リスト選択
    hanso2_frame = ttk.LabelFrame(root_window, text="版素コード（特殊）-リスト選択")
    hanso2_frame.pack(pady=10, fill="x", padx=10)
    hanso4_var = tk.StringVar()
    ttk.Label(hanso2_frame).pack(side="left", padx=5)
    ttk.Combobox(hanso2_frame, width=25, background="#fafad2", 
                 textvariable=hanso4_var, 
                 values=["", "KRS-A1：完全支給データ","NTD-A1：色校のみ","OLS-A1：見積書あり"]).pack(side="left", padx=10)

    # 連絡事項フリーテキスト
    announce_text_frame = ttk.LabelFrame(root_window, text="その他連絡事項")
    announce_text_frame.pack(pady=10, fill="both", padx=10, expand=True)
    announce_text = tk.Text(announce_text_frame, wrap="word", height=1)
    announce_text.pack(fill="both", padx=10, pady=5, expand=True)

    # 実行ボタン
    execute_button = ttk.Button(root_window, text="実行", width=50, command=on_execute)
    execute_button.pack(pady=20)

    root_window.mainloop()

if __name__ == "__main__":
    tenbun_order_gui()
