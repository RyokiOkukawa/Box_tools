#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Administrator
#
# Created:     28/05/2020
# Copyright:   (c) Administrator 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Administrator
#
# Created:     28/05/2020
# Copyright:   (c) Administrator 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#import モジュール

import glob
import codecs
from lxml import html

#データ取得用変数リスト
src_path = []
dst_path = []
dst_user = []

comparative_start_datetime = []
comparative_time = []
migration_start_datetime = []
migration_time = []
total_size = []
migration_size = []
task = []
error = []
end_datetime = []
folder_total = []
folder_to_dst = []
folder_to_src = []
folder_skip = []
folder_deleted = []
folder_miss = []
file_total = []
file_to_dst = []
file_to_src = []
file_skip = []
file_deleted = []
file_miss = []
size_total = []
size_to_dst = []
size_to_src = []
size_skip = []
size_deleted = []
size_miss = []

#ソースパス

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[197].decode('utf_8'))

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        src_path.append(memo)

#移行先のユーザー
for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[191].decode('utf_8'))

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        dst_user.append(memo)

#移行先のファイルパス取得
for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[200].decode('utf_8'))

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        dst_path.append(memo)

#比較開始時間

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[223])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        comparative_start_datetime.append(memo)

#比較時間

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[229])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        comparative_time.append(memo)

#移行開始時間

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[238])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        migration_start_datetime.append(memo)

#比較時間

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[244])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        migration_time.append(memo)

#合計サイズ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[253])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        total_size.append(memo)

#移行サイズ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[259])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        migration_size.append(memo)

#タスク数

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[268])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        task.append(memo)

#エラー

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[274])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        error.append(memo)

#終了時刻

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[282])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        end_datetime.append(memo)

#フォルダ 合計

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[349])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_total.append(memo)

#フォルダ 移行先へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[352])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_to_dst.append(memo)

#フォルダ 移行元へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[355])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_to_src.append(memo)

#フォルダ スキップ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[358])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_skip.append(memo)

#フォルダ 削除済み

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[361])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_deleted.append(memo)

#フォルダ 失敗

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[364])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        folder_miss.append(memo)

#ファイル 合計

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[373])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_total.append(memo)

#ファイル 移行先へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[376])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_to_dst.append(memo)

#ファイル 移行元へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[379])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_to_src.append(memo)

#ファイル スキップ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[382])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_skip.append(memo)

#ファイル 削除済み

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[385])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_deleted.append(memo)

#ファイル 失敗

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[388])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        file_miss.append(memo)

#サイズ 合計

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[397])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_total.append(memo)

#サイズ 移行先へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[400])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_to_dst.append(memo)

#サイズ 移行元へ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[403])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_to_src.append(memo)

#サイズ スキップ

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[406])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_skip.append(memo)

#サイズ 削除済み

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[409])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_deleted.append(memo)

#サイズ 失敗

for file in glob.glob(r"D:\dis_box_Tools\box_summery\work5\*.html"):
    with open(file, mode='rb') as g:
        t = html.fromstring(g.readlines()[412])

        remove_tags = ('.//style', './/script', './/noscript')
        for remove_tag in remove_tags:
            for tag in t.findall(remove_tag):
                tag.drop_tree()

        memo = t.text_content().strip()
        size_miss.append(memo)

#取得情報のリスト化

summary = [src_path,
           dst_user,
           dst_path,
           comparative_start_datetime,
           comparative_time,
           migration_start_datetime,
           migration_time,
           total_size,
           migration_size,
           task,
           error,
           end_datetime,
           folder_total,
           folder_to_dst,
           folder_to_src,
           folder_skip,
           folder_deleted,
           folder_miss,
           file_total,
           file_to_dst,
           file_to_src,
           file_skip,
           file_deleted,
           file_miss,
           size_total,
           size_to_dst,
           size_to_src,
           size_skip,
           size_deleted,
           size_miss
           ]

#変数summaryをパイプ区切りでzip化

summary_zip = list(map(list, zip(*summary)))
for i in summary_zip:
    summary_b = '|'.join(i)
    print(summary_b, file = codecs.open(r"D:\dis_box_Tools\box_summery\DMB\summary.txt", 'a', 'cp932'))


#不要文字列の削除

with open(r"D:\dis_box_Tools\box_summery\DMB\summary.txt", encoding = 'cp932') as f:
    data_lines = f.read()

data_lines = data_lines.replace("パス  ", "")
data_lines = data_lines.replace("アカウント ", "")
data_lines = data_lines.replace("パス ", "")

with open(r"D:\dis_box_Tools\box_summery\DMB\summary.txt", 'w', encoding = 'cp932') as f:
    f.write(data_lines)
