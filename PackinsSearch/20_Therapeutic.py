# PMDAサイトから薬効分類ごとの検索で全点データを取得する試み
# 2022.10.25

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

import bs4
from bs4 import BeautifulSoup
from bs4 import Tag

import os
import subprocess
import re
import urllib.request
import time
import datetime
import sys
import math
import platform

#----- 初期処理 -----
print("### データ取得開始 ###")
tb_total = 0  #総数カウント用
## ----- カレントディレクトリ -----
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
## ----- 出力ファイル設定（旧ファイルが残っている場合は削除）-----
global opt_fil
opt_fil = 'pmda_tenbun_all.csv'
if os.path.exists(opt_fil):
    os.remove(opt_fil)
# -----
def editstrings(str_in): #文字列処理---ファイル名に使えない文字を変換または削除
    str_in = re.sub('^_|\&amp;|[\\\/\:\;\=\*\?\"\s\<\>\|\,\.\%®µçß]',"",str_in)
    str_in = re.sub('(\u207b|\u2212|⁻|−)','-',str_in)
    str_in = re.sub('[àâä]','a',str_in)
    str_in = re.sub('[èéêë]','e',str_in)
    str_in = re.sub('[ôö]','o',str_in)
    str_in = re.sub('[ïî]','i',str_in)
    str_in = re.sub('[ûùü]','u',str_in)            
    str_in = re.sub('(µ|\u00b5)','μ',str_in)
    str_in = re.sub('㎍','μg',str_in)
    str_in = re.sub('Ä','A',str_in)
    str_in = re.sub('Ö','O',str_in)
    str_in = re.sub('Ü','U',str_in)
    str_in = re.sub('\u2f00','一',str_in)
    str_in = re.sub('\u2f04','乙',str_in)
    str_in = re.sub('\u2f06','二',str_in)
    str_in = re.sub('\u2f08','人',str_in)
    str_in = re.sub('\u2f0a','入',str_in)
    str_in = re.sub('\u2f0b','八',str_in)
    str_in = re.sub('\u2f11','刀',str_in)
    str_in = re.sub('\u2f12','力',str_in)
    str_in = re.sub('\u2f17','十',str_in)
    str_in = re.sub('\u2f1e','口',str_in)
    str_in = re.sub('\u2f1f','土',str_in)
    str_in = re.sub('\u2f20','士',str_in)
    str_in = re.sub('\u2f23','夕',str_in)
    str_in = re.sub('\u2f24','大',str_in)
    str_in = re.sub('\u2f25','女',str_in)
    str_in = re.sub('\u2f26','子',str_in)
    str_in = re.sub('\u2f28','寸',str_in)
    str_in = re.sub('\u2f29','小',str_in)
    str_in = re.sub('\u2f2d','山',str_in)
    str_in = re.sub('\u2f2f','工',str_in)
    str_in = re.sub('\u2f30','己',str_in)
    str_in = re.sub('\u2f31','巾',str_in)
    str_in = re.sub('\u2f32','干',str_in)
    str_in = re.sub('\u2f38','弓',str_in)
    str_in = re.sub('\u2f3c','心',str_in)
    str_in = re.sub('\u2f3e','戸',str_in)
    str_in = re.sub('\u2f3f','手',str_in)
    str_in = re.sub('\u2f40','支',str_in)
    str_in = re.sub('\u2f42','分',str_in)
    str_in = re.sub('\u2f43','斗',str_in)
    str_in = re.sub('\u2f44','斤',str_in)
    str_in = re.sub('\u2f45','方',str_in)
    str_in = re.sub('\u2f47','日',str_in)
    str_in = re.sub('\u2f48','曰',str_in)
    str_in = re.sub('\u2f49','月',str_in)
    str_in = re.sub('\u2f4a','木',str_in)
    str_in = re.sub('\u2f4b','欠',str_in)
    str_in = re.sub('\u2f4c','止',str_in)
    str_in = re.sub('\u2f4f','母',str_in)
    str_in = re.sub('\u2f50','比',str_in)
    str_in = re.sub('\u2f51','毛',str_in)
    str_in = re.sub('\u2f52','氏',str_in)
    str_in = re.sub('\u2f54','水',str_in)
    str_in = re.sub('\u2f55','火',str_in)
    str_in = re.sub('\u2f56','爪',str_in)
    str_in = re.sub('\u2f57','父',str_in)
    str_in = re.sub('\u2f5a','片',str_in)
    str_in = re.sub('\u2f5b','牙',str_in)
    str_in = re.sub('\u2f5c','牛',str_in)
    str_in = re.sub('\u2f5d','犬',str_in)
    str_in = re.sub('\u2f5e','玄',str_in)
    str_in = re.sub('\u2f5f','玉',str_in)
    str_in = re.sub('\u2f60','瓜',str_in)
    str_in = re.sub('\u2f61','瓦',str_in)
    str_in = re.sub('\u2f62','甘',str_in)
    str_in = re.sub('\u2f63','生',str_in)
    str_in = re.sub('\u2f64','用',str_in)
    str_in = re.sub('\u2f65','田',str_in)
    str_in = re.sub('\u2f66','疋',str_in)
    str_in = re.sub('\u2f69','白',str_in)
    str_in = re.sub('\u2f6a','皮',str_in)
    str_in = re.sub('\u2f6b','皿',str_in)
    str_in = re.sub('\u2f6c','目',str_in)
    str_in = re.sub('\u2f6d','矛',str_in)
    str_in = re.sub('\u2f6e','矢',str_in)
    str_in = re.sub('\u2f6f','石',str_in)
    str_in = re.sub('\u2f70','示',str_in)
    str_in = re.sub('\u2f72','禾',str_in)
    str_in = re.sub('\u2f73','穴',str_in)
    str_in = re.sub('\u2f74','立',str_in)
    str_in = re.sub('\u2f75','竹',str_in)
    str_in = re.sub('\u2f76','米',str_in)
    str_in = re.sub('\u2f77','糸',str_in)
    str_in = re.sub('\u2f78','缶',str_in)
    str_in = re.sub('\u2f7a','羊',str_in)
    str_in = re.sub('\u2f7b','羽',str_in)
    str_in = re.sub('\u2f7c','老',str_in)
    str_in = re.sub('\u2f7d','而',str_in)
    str_in = re.sub('\u2f7f','耳',str_in)
    str_in = re.sub('\u2f81','肉',str_in)
    str_in = re.sub('\u2f82','臣',str_in)
    str_in = re.sub('\u2f83','自',str_in)
    str_in = re.sub('\u2f84','至',str_in)
    str_in = re.sub('\u2f85','臼',str_in)
    str_in = re.sub('\u2f86','舌',str_in)
    str_in = re.sub('\u2f87','舛',str_in)
    str_in = re.sub('\u2f88','舟',str_in)
    str_in = re.sub('\u2f8a','色',str_in)
    str_in = re.sub('\u2f90','衣',str_in)
    str_in = re.sub('\u2f92','見',str_in)
    str_in = re.sub('\u2f93','角',str_in)
    str_in = re.sub('\u2f94','言',str_in)
    str_in = re.sub('\u2f95','谷',str_in)
    str_in = re.sub('\u2f96','豆',str_in)
    str_in = re.sub('\u2f99','貝',str_in)
    str_in = re.sub('\u2f9a','赤',str_in)
    str_in = re.sub('\u2f9b','走',str_in)
    str_in = re.sub('\u2f9c','足',str_in)
    str_in = re.sub('\u2f9d','身',str_in)
    str_in = re.sub('\u2f9e','車',str_in)
    str_in = re.sub('\u2f9f','辛',str_in)
    str_in = re.sub('\u2fa0','辰',str_in)
    str_in = re.sub('\u2fa2','邑',str_in)
    str_in = re.sub('\u2fa3','酉',str_in)
    str_in = re.sub('\u2fa4','采',str_in)
    str_in = re.sub('\u2fa5','里',str_in)
    str_in = re.sub('\u2fa6','金',str_in)
    str_in = re.sub('\u2fa7','長',str_in)
    str_in = re.sub('\u2fa8','門',str_in)
    str_in = re.sub('\u2fa9','阜',str_in)
    str_in = re.sub('\u2fac','雨',str_in)
    str_in = re.sub('\u2fae','非',str_in)
    str_in = re.sub('\u2faf','面',str_in)
    str_in = re.sub('\u2fb0','革',str_in)
    str_in = re.sub('\u2fb3','音',str_in)
    str_in = re.sub('\u2fb4','頁',str_in)
    str_in = re.sub('\u2fb5','風',str_in)
    str_in = re.sub('\u2fb6','飛',str_in)
    str_in = re.sub('\u2fb7','食',str_in)
    str_in = re.sub('\u2fb8','首',str_in)
    str_in = re.sub('\u2fb9','香',str_in)
    str_in = re.sub('\u2fba','馬',str_in)
    str_in = re.sub('\u2fbb','骨',str_in)
    str_in = re.sub('\u2fbc','高',str_in)
    str_in = re.sub('\u2fc1','鬼',str_in)
    str_in = re.sub('\u2fc2','魚',str_in)
    str_in = re.sub('\u2fc3','鳥',str_in)
    str_in = re.sub('\u2fc5','鹿',str_in)
    str_in = re.sub('\u2fc7','朝',str_in)
    str_in = re.sub('\u2fca','黒',str_in)
    str_in = re.sub('\u2fce','鼓',str_in)
    str_in = re.sub('\u2fcf','鼠',str_in)
    str_in = re.sub('\u2fd0','鼻',str_in)
    str_in = re.sub('\u2fd1','齊',str_in)
    str_in = re.sub('\u2fd3','龍',str_in)
    return(str_in)


def PackinsCount(class_code, class_name):
    driver.get('https://www.info.pmda.go.jp/psearch/html/menu_tenpu_kensaku.html')
    #条件消去ボタンをクリック
    driver.find_element(By.CSS_SELECTOR, "input[value='条件消去']").click()
    #その他検索条件指定ボタンをクリック
    driver.find_element(By.ID, "visibleBtn").click()
    driver.find_element(By.XPATH, "//input[@id='targetBothWithSgmlItem']").click() # 「新旧様式」を選択
    # 薬効分類名を選択
    KEYWORD = class_code # パラメータで受け取った値は分類コード
    search_box = driver.find_element(By.NAME, "effect")
    select = Select(search_box)
    select.select_by_value(KEYWORD)
    COUNT = "100" # 表示件数は最大値の100
    count_box = driver.find_element(By.NAME, "count")
    select2 = Select(count_box)
    select2.select_by_value(COUNT)
    # 生薬用 例外処理（1000件over）
    if class_code == '51' or class_code == '510':
        KEYWORD2 = "う"
        search_box = driver.find_element(By.NAME, "keyword1")
        search_box.send_keys(KEYWORD2)
        if len(class_code) == 2:
            search_type = "not"
        else:
            search_type = "and"
        type_box = driver.find_element(By.NAME, "type1")
        select3 = Select(type_box)
        select3.select_by_value(search_type)
    
    driver.find_element(By.CSS_SELECTOR, "input[value='検索実行']").click()
    #time.sleep(3)

    driver.switch_to.window(driver.window_handles[1])
    result_pmda1 = driver.page_source
    time.sleep(2)
    soup = BeautifulSoup(result_pmda1, "html.parser")
    Num_class = Numofclass(soup)
    print(class_name, Num_class)
    htm_opt = open(opt_fil, 'a', encoding='UTF-8')
    if Num_class == None:
        Num_class = '0'
        htm_opt.write('■' + class_name + "：" + Num_class + '\n')
        return(0)
    elif int(Num_class) > 1000:
        htm_opt.write('■検索上限超■' + class_name + "：" + Num_class + '\n')
        return(0)
    else:
        htm_opt.write('■' + class_name + "：" + Num_class + '\n')
    htm_opt.close()
    
    all_pages = GetAllPage(soup)
    for elm_page in all_pages:
        PickoutProductinfo(elm_page)
    
    return(Num_class)

# 薬効分類ごとの検索ヒット数
def Numofclass(soup):
    for elm_pmda in soup.find_all('font'):
        num_pmda = elm_pmda.get_text()
        if re.compile('^\d+$').search(num_pmda):
            return(num_pmda)
# 対象頁を一通り取得してリストに格納
def GetAllPage(soup):
    soup_pages = [soup]
    crnt_src = soup
    flg_nextpage = 1
    while flg_nextpage == 1:
        for elm_pmda2 in crnt_src.find_all('a'):
            flg_nextpage = 0
            if elm_pmda2.get_text() == '次へ':
                flg_nextpage = 1
                try:
                    driver.find_element(By.LINK_TEXT, '次へ').click()
                    time.sleep(3)
                    result_pmda2 = driver.page_source
                    crnt_src = BeautifulSoup(result_pmda2, "html.parser")
                    soup_pages.append(crnt_src)
                except:
                    break
    return(soup_pages)

def PickoutProductinfo(soup):
    htm_opt = open(opt_fil, 'a', encoding='UTF-8')
    soup2 = []
    sgml_mark = ''
    for elm_pmda1 in soup.find_all(['a','tr']):
        if elm_pmda1.name == 'a':
            prdt_name = elm_pmda1.get_text()
            prdt_name = re.sub('(\ |\t|　|\*|※|＊|®|,|、|，|\/|\u00a0|\u0020|[\r\n]+)','',prdt_name)
            prdt_id = elm_pmda1.get('href')
            prdt_id = re.search("\d[0-9A-Z]{11}_._[0-9A-Z]{2}", prdt_id)
            if prdt_id != None:
                prdt_id_str = prdt_id.group()
        elif elm_pmda1.name == 'tr' and elm_pmda1.get('bgcolor') == "lightblue":
            upd_date = ''
            #print(elm_pmda1)
            cop_list = []
            for detail_pmda in elm_pmda1.find_all(['font']):
                if re.compile("更新日").search(detail_pmda.get_text()):
                    upd_date = re.sub("更新日：", "", detail_pmda.get_text())
                    detail_pmda.decompose()
                else:
                    detail_pmda.decompose()
                #print("●", elm_pmda1.get_text())
                #print("●", upd_date)
            for detail_pmda in elm_pmda1.find_all(['dd']):
                cop_name = detail_pmda.get_text()
                #cop_name = re.sub('(^(.*)?／|\ |\t|　|\*|※|＊|®|,|、|，|\/|\u00a0|\u0020|(製造)?販売元|[\r\n]+)', '', cop_name)
                if cop_name != '':
                    cop_list.append(cop_name)
            #print("◆", prdt_name, prdt_id_str,cop_list, upd_date)
            out_data = sgml_mark + "\t" + prdt_name + "\t" + prdt_id_str + "\t" + upd_date + "\t"
            sgml_mark = ''
            for elm_copname in cop_list:
                elm_copname = editstrings(elm_copname)
                out_data += elm_copname + "\t"
            out_data_ed = editstrings(out_data)
            htm_opt.write(out_data + '\n')
        elif elm_pmda1.name == 'tr':
            for cls_sgml in elm_pmda1.find_all(['span']):
                sgml_mark = cls_sgml.get_text()
                #input(sgml_mark)

    htm_opt.close()
    return(soup2)

#----------------------------------------------------------------------    
options = Options()
#options.add_argument('--headless') # ブラウザ非表示

global driver
driver = webdriver.Firefox(options=options)

class_list = ["735：Dummy",
"111：全身麻酔剤","112：催眠鎮静剤，抗不安剤","113：抗てんかん剤","114：解熱鎮痛消炎剤",
"115：興奮剤，覚醒剤","116：抗パーキンソン剤","117：精神神経用剤","118：総合感冒剤",
"119：その他の中枢神経系用薬","121：局所麻酔剤","122：骨格筋弛緩剤","123：自律神経剤",
"124：鎮けい剤","125：発汗剤，止汗剤","129：その他の末梢神経系用薬","131：眼科用剤",
"132：耳鼻科用剤","133：鎮暈剤","139：その他の感覚器官用薬","19：その他の神経系及び感覚器官用医薬品",
"211：強心剤","212：不整脈用剤","213：利尿剤","214：血圧降下剤","215：血管補強剤","216：血管収縮剤",
"217：血管拡張剤","218：高脂血症用剤","219：その他の循環器官用薬","221：呼吸促進剤","222：鎮咳剤",
"223：去たん剤","224：鎮咳去たん剤","225：気管支拡張剤","226：含嗽剤","229：その他の呼吸器官用薬",
"231：止しゃ剤，整腸剤","232：消化性潰瘍用剤","233：健胃消化剤","234：制酸剤","235：下剤，浣腸剤",
"236：利胆剤","237：複合胃腸剤","239：その他の消化器官用薬","241：脳下垂体ホルモン剤","242：唾液腺ホルモン剤",
"243：甲状腺，副甲状腺ホルモン剤","244：たん白同化ステロイド剤","245：副腎ホルモン剤","246：男性ホルモン剤",
"247：卵胞ホルモン及び黄体ホルモン剤","248：混合ホルモン剤","249：その他のホルモン剤（抗ホルモン剤を含む）",
"251：泌尿器官用剤","252：生殖器官用剤（性病予防剤を含む。）","253：子宮収縮剤","254：避妊剤",
"255：痔疾用剤","259：その他の泌尿生殖器官及び肛門用薬","261：外皮用殺菌消毒剤","262：創傷保護剤",
"263：化膿性疾患用剤","264：鎮痛，鎮痒，収歛，消炎剤","265：寄生性皮ふ疾患用剤",
"266：皮ふ軟化剤（腐しょく剤を含む。）","267：毛髪用剤（発毛剤，脱毛剤，染毛剤，養毛剤","268：浴剤",
"269：その他の外皮用薬","271：歯科用局所麻酔剤","272：歯髄失活剤",
"273：歯科用鎮痛鎮静剤（根管及び齲窩消毒剤を含","274：歯髄乾屍剤（根管充填剤を含む。）",
"275：歯髄覆たく剤","276：歯科用抗生物質製剤","279：その他の歯科口腔用薬","290：その他の個々の器官系用医薬品",
"311：ビタミンＡ及びＤ剤","312：ビタミンＢ１剤","313：ビタミンＢ剤（ビタミンＢ１剤を除く。）",
"314：ビタミンＣ剤","315：ビタミンＥ剤","316：ビタミンＫ剤",
"317：混合ビタミン剤（ビタミンＡ・Ｄ混合製剤を除く）","319：その他のビタミン剤","321：カルシウム剤",
"322：無機質製剤","323：糖類剤","324：有機酸製剤","325：たん白アミノ酸製剤","326：臓器製剤",
"327：乳幼児用剤","329：その他の滋養強壮薬","331：血液代用剤","332：止血剤","333：血液凝固阻止剤",
"339：その他の血液・体液用薬","341：人工腎臓透析用剤","342：腹膜透析用剤","349：その他の人工透析用薬",
"391：肝臓疾患用剤","392：解毒剤","393：習慣性中毒用剤","394：痛風治療剤","395：酵素製剤","396：糖尿病用剤",
"397：総合代謝性製剤","399：他に分類されない代謝性医薬品","411：クロロフィル製剤","412：色素製剤",
"419：その他の細胞賦活用薬","421：アルキル化剤","422：代謝拮抗剤","423：抗腫瘍性抗生物質製剤",
"424：抗腫瘍性植物成分製剤","429：その他の腫瘍用薬","430：放射性医薬品","441：抗ヒスタミン剤",
"442：刺激療法剤","443：非特異性免疫原製剤","449：その他のアレルギー用薬","490：その他の組織細胞機能用医薬品",
"51：生薬","510：生薬","520：漢方製剤","590：その他の生薬及び漢方処方に基づく医薬品",
"611：主としてグラム陽性菌に作用するもの","612：主としてグラム陰性菌に作用するもの",
"613：主としてグラム陽性・陰性菌に作用するもの","614：主としてグラム陽性菌，マイコプラズマに作用するもの",
"615：主としてグラム陽性・陰性菌，リケッチア，クラミジアに作用するもの","616：主として抗酸菌に作用するもの",
"617：主としてカビに作用するもの","619：その他の抗生物質製剤（複合抗生物質製剤を含む）","621：サルファ剤",
"622：抗結核剤","623：抗ハンセン病剤","624：合成抗菌剤","625：抗ウイルス剤","629：その他の化学療法剤",
"631：ワクチン類","632：毒素及びトキソイド類","633：抗毒素類及び抗レプトスピラ血清類","634：血液製剤類",
"635：生物学的試験用製剤類","636：混合生物学的製剤","639：その他の生物学的製剤","641：抗原虫剤",
"642：駆虫剤","649：その他の寄生動物用薬","690：その他の病原生物に対する医薬品","711：賦形剤",
"712：軟膏基剤","713：溶解剤","714：矯味，矯臭，着色剤","715：乳化剤",
"719：その他の調剤用薬","721：Ｘ線造影剤","722：機能検査用試薬","729：その他の診断用薬",
"731：防腐剤","732：防疫用殺菌消毒剤","733：防虫剤","734：殺虫剤","735：殺そ剤",
"739：その他の公衆衛生用薬","791：ばん創こう","799：他に分類されない治療を主目的としない医薬品",
"811：アヘンアルカロイド系麻薬","812：コカアルカロイド系製剤","819：その他のアルカロイド系麻薬（天然麻薬）",
"821：合成麻薬"
]
for elm_cls in class_list:
    elm_cls_s = elm_cls.split('：')
    cls_code = elm_cls_s[0]
    cls_name = elm_cls_s[1]
    time.sleep(2)
    cls_num = PackinsCount(cls_code, cls_name)
    tb_total += int(cls_num)
    print(tb_total)
print('---Finished---')