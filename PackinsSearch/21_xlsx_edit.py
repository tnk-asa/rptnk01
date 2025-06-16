#2024.01.24 ChatGPTで作成
#添付文書全件データを編集
#　製造販売元列を追加
#　Excelで使用可能な形式で出力

import csv
import re
import os

#.pyがあるディレクトリをカレントディレクトリに
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 入力ファイル
input_file_path = r"pmda_tenbun_all.csv"
# 出力ファイル
output_file_path = r"rep_pmda_tenbun.csv"
# ヘッダ行ラベル
header_labels = ["薬効分類", "新旧", "販売名", "添付文書番号", "更新日", "業者１", "業者２", "業者３", "製造販売元"]

# 出力ファイルが既に存在する場合は上書きする
with open(output_file_path, mode='w', encoding='utf-8-sig', newline='') as output_file:
    # CSVライターを作成
    csv_writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
    # ヘッダ行を書き込む
    csv_writer.writerow(header_labels)
    # 保持する文字列を初期化
    preserved_string = ''
    # 薬効分類のフラグ
    manufacturer_found = False

    # 入力ファイルからデータを読み込み、編集して書き込む
    with open(input_file_path, mode='r', encoding='utf-8-sig', newline='') as input_file:
        csv_reader = csv.reader(input_file, delimiter='\t')

        for row in csv_reader:
            if row and row[0].startswith('■'):
                # ■で始まる行の処理
                preserved_string = row[0][1:]
                manufacturer_found = False  # ■で始まる行があったらフラグをリセット
            else:
                # 先頭が■ではない行が出現した場合、保持した文字列を挿入
                if preserved_string:
                    row.insert(0, preserved_string)

                # 行の列数がヘッダと一致していない場合は整形
                if len(row) < 9:
                    row += [''] * (9 - len(row))
                elif len(row) > 9:
                    row = row[:9]

                # 6列目、7列目、8列目の文字列に指定された正規表現が含まれていれば、その内容を9列目に入力
                out_chr = ""
                for i in range(5, 8):
                    if re.search(r'(選任|外国|医薬品)*製造(販売|発売|販)?元?', str(row[i])):
                        out_chr = re.sub('^.*?／','',row[i])
                        out_chr = re.sub('(製造販売(業者)?|（輸(入元）)?|本注意|[※＊、，])','',out_chr)
                        ptn_zenkaku = re.compile(r'[０-９Ａ-Ｚａ-ｚ]')
                        out_chr = re.sub(ptn_zenkaku, lambda x: chr(ord(x.group(0)) - 0xFEE0), out_chr)
                        row[8] = out_chr
                        break
                    elif i == 7:
                        out_chr = row[5]
                        out_chr = re.sub('^.*?／','',out_chr)
                        out_chr = re.sub('(製造販売(業者)?|（輸(入元）)?|本注意|[※＊、，])','',out_chr)
                        ptn_zenkaku = re.compile(r'[０-９Ａ-Ｚａ-ｚ]')
                        out_chr = re.sub(ptn_zenkaku, lambda x: chr(ord(x.group(0)) - 0xFEE0), out_chr)
                        row[8] = out_chr
                csv_writer.writerow(row)

print("編集が完了しました。")
