"""
pipで以下のモジュールが必要になります。
py -m pip install -U pip をcmdで実行後

pip install pandas
pip install beautifulsoup4
"""

import glob
import os
import pandas as pd  # pipでのインストールが必要
import bs4  # pipでのインストールが必要
from lxml import html

excel_index = []  # 行番号を格納するリスト
src_path = []  # パスを格納するリスト
out_path_list = []  # 出力するパスリスト
out_error_list = []  # 出力するエラーリスト
out_error_path_list = []  # 出力するエラーパスリスト
out_time_list = []  # 出力するエラータイムリスト
out_size_list = []  # 出力するサイズリスト

'''
パスを取得するループ
ここから
'''
for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[197].decode('utf_8'))
        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        src_path.append(memo)
'''
パスを取得するループ
ここまで
'''
'''
エラー情報を取得するループ
ここから
'''
files = os.listdir(path=r'D:\dis_box_Tools\box_summery\work5')  # 対象フォルダ内のファイル名を全取得 ここでhtmlファイルのあるフォルダを指定
i = 0  # パスを格納したリスト(src_path)の要素番号を指定する変数
for html_name in files:  # 取得したファイル名を一つずつ取り出す
    soup = bs4.BeautifulSoup(open('D:\\dis_box_Tools\\box_summery\work5\\' + html_name, encoding="utf-8"),
                             'html.parser')  # htmlファイルの全要素取得
    error_elem = soup.find_all(class_="cell_large")  # htmlの中からエラー内容を格納
    time_elems = soup.find_all(class_="cell")[45::3]  # htmlの中からエラー時間を格納
    size_elems = soup.find_all(class_="cell")[46::3]  # htmlの中からサイズを格納

    for value in error_elem:  # エラー内容とパスをリストへ格納するループ
        html_elems_list = value.text  # エラー内容をリストへ格納
        t = html_elems_list.strip()  # エラー内容の前後にある空白を取り除く

        if '\uffff' in t:  # ユニコード回避
            t = t.replace('\uffff', '')

        if '失敗' in t:  # 格納する文字列は失敗と書いてあるものだけで良さそうなので絞る
            path = src_path[i]  # パスを変数へ格納
            out_path_list.append(path)  # 出力用のリストへパスを格納
            out_error_list.append(t)  # 出力用のリストへエラー内容を格納
            # print(t)
        else:
            out_error_path_list.append(t)
            # print(t)
    i += 1  # パスの要素番号をインクリメント

    for value in time_elems:  # エラー発生時間を格納するループ
        cell_elems = value.text  # エラー発生時間を変数へ格納
        t = cell_elems.strip()  # エラー発生時間の前後の空白を取り除く
        out_time_list.append(t)  # エラー発生時間を出力用リストへ格納

    for value in size_elems:  # サイズを格納するループ
        cell_elems = value.text  # サイズを変数へ格納
        t = cell_elems.strip()  # サイズの前後の空白を取り除く
        out_size_list.append(t)  # サイズを出力用リストへ格納
'''
エラー情報を取得するループ
ここまで
'''
'''
出力データのフレーム作成
ここから
'''
for i in range(len(out_path_list)):  # 行番号を要素分取得するループ分
    excel_index.append(i + 1)

out_list = [[0 for j in range(5)] for s in range(len(out_path_list))]  # 多次元配列を作成
for i in range(len(out_list)):  # 多次元配列に要素を詰め込んでいくここはkey値を指定
    for j in range(len(out_list[i])):  # value値を格納
        if j == 0:  # 各要素の住所に入れる条件分岐
            out_list[i][j] = out_path_list[i]  # パスの住所は0
        if j == 1:
            out_list[i][j] = out_error_path_list[i]  # エラーパスの住所は1
        if j == 2:
            out_list[i][j] = out_error_list[i]  # エラー内容の住所は2
        if j == 3:
            out_list[i][j] = out_time_list[i]  # 時間の住所は3
        if j == 4:
            out_list[i][j] = out_size_list[i]  # サイズの住所は4

df = pd.DataFrame(out_list, index=excel_index, columns=['パス', 'エラーパス', 'エラー', 'エラー発生時間', 'サイズ'])  # エクセルのデータフレーム作成
df.to_excel(r'D:\dis_box_Tools\box_summery\summary\summary.xlsx')  # エクセルファイル作成 & データを書き出し ←ここにの保存先とエクセル名を記入
'''
出力データのフレーム作成
ここまで

作成者：奥川
最終更新:2020/08/03
'''
