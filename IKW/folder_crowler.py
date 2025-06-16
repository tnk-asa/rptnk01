import os
import re

os.chdir(os.path.dirname(os.path.abspath(__file__))) #.pyがあるディレクトリをカレントディレクトリに

def get_matching_files(folder_path, extensions):
    result_list = []

    # 正規表現のパターン（例：123456_AB12CD34EF56_Z）
    pattern = re.compile(r'\d{6}_[A-Za-z0-9]{16}_[A-Z].*$')

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            name, ext = os.path.splitext(file)
            if ext.lower() in [e.lower() for e in extensions] and pattern.match(name):
                result_list.append(os.path.join(root, file))

    return result_list

from bs4 import BeautifulSoup
import os

#def extract_tags_to_txt(xml_file_path, tag_list=["Item", "Detail"], output_path=None):
def extract_tags_to_txt(xml_file_path, tag_list, output_path=None):
    """
    XMLファイルから特定タグを抽出して、テキストに追記保存
    Parameters:
        xml_file_path (str): 入力するXMLファイルのパス
        tag_list (list): タグ名リスト（例：["Item", "Detail"]）
        output_path (str): 出力ファイルのパスなければカレントに extracted.txt
    """
    try:
        if output_path is None:
            output_path = os.path.join(os.getcwd(), "extracted.txt")

        with open(xml_file_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "xml")

        base_name = os.path.basename(xml_file_path)

        # 出力内容リストの初期化（ファイル名から始まるよ）
        extracted_texts = [f"===== {base_name} ====="]

        for tag in tag_list:
            for element in soup.find_all(tag):
                raw = str(element)  # タグごと文字列化しちゃう
                extracted_texts.append(raw)

        # 書き出しは追記モード！
        with open(output_path, "a", encoding="utf-8") as out_file:
            for line in extracted_texts:
                out_file.write(line + "\n")

        print(f"{base_name} からタグごと抽出して {output_path} に追記完了")

    except Exception as e:
        print(f"エラー発生！内容：{e}")



#--------------------------------------------------
folder = r"C:\\Users\\P0244\\Downloads\\ikw_xml_all"
exts = ['.txt', '.xml']
files = get_matching_files(folder, exts)

for f in files:
    print(f)
    extract_tags_to_txt(f, ["ApprovalBrandName", "Classification", "InfoIndicationsOrEfficacy"]) # 対象タグ