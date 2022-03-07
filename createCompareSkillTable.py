"""
比較技能分析表作成コード
Python3.8.8
①anacondaをインストールする or　pandas(1.2.5) ,pywin32(227)をインストール
 pip install pywin32==227
 pip install pandas==1.2.5

②発言録まとめにkoutei_id,sagyo_id,koudou_idを振る
　被験者間で隣に配置したい行に同じidを振る

③このファイルと同じ階層にimages, videos, 発言録まとめファイルを置く
imagesに画像、videosに動画を入れる（image列、video列とファイル名が）

④以下3つを入力
1. file_name: 発言録まとめのファイル名
2. img_size: 貼り付ける画像のサイズ（高さを指定、横幅はアスペクト比固定で計算）
3. save_name: 保存ファイル名, (output/○○_y-m-d-h-m-s.xlsxで保存される)

⑤不具合がないか確認する（#NAが入ったりしてしまいます。）
"""

#以下3つを入力++++++++++++++++++++++++++++++++++++++++++
#ファイルの名前
file_name = "中間報告用発言録まとめ.xlsx"

#画像のheight(cm)
img_size = 5

#保存ファイル名
save_name = "C-Nexco"
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++


"""
コードの流れ概要
IDsというDFでkoutei_id, sagyo_id, koudou_idをuniqueで持つ
IDsを1行ずつループ
　被験者数ループ
　  koutei_id, sagyo_id, koudou_idが全て一致する部分を各被験者の発言録まとめから抜き出す
　　1人分の動画貼る
　　1人分の画像貼る
　　↑人数分繰り返し
　人数分のmini header挿入（画像数に応じて被験者間で終わる行が変わるため画像等挿入後に実行）
　被験者数ループ
　　1人分行動の内容挿入
　　↑人数分繰り返し
　コード上に戻り行動/作業が終了のタイミングでセル結合
最終行のセル結合
体裁を整える
保存
"""

import datetime
import pandas as pd
import win32com.client
import os

#列名の変更等に耐えられるよう定数を設定しましたが、完成を急ぐあまり、機能していないです。
#もし発言録まとめの列名が変更された時のための定数
KOUTEI_ID = "koutei_ID"
SAGYO_ID = "sagyo_ID"
KOUDOU_ID = "koudou_ID"
KOUTEI = "工程"
SAGYO = "主な作業内容"
#列を指定する際にループで増加させるインデックスの増加定数
INDEX_PLUS = 5

#発言録まとめの読み込み
sheet_names = pd.ExcelFile(file_name).sheet_names
excels = pd.read_excel(file_name, sheet_name=sheet_names)

#ExcelAppの起動
xlApp = win32com.client.Dispatch("Excel.Application")

#IDs DFの作成（idのマトリクス）
#koutei_id, sagyo_id, koudou_id, 工程名, 作業名 を列にもつDF
colmuns = [KOUTEI_ID, SAGYO_ID, KOUDOU_ID, KOUTEI, SAGYO]
IDs = pd.DataFrame(index=[], columns=colmuns)
for sheet_name in sheet_names:
    excel = excels[sheet_name]
    excel[KOUTEI_ID] = excel[KOUTEI_ID].fillna(method="ffill").astype(int)
    excel[SAGYO_ID] = excel[SAGYO_ID].fillna(method="ffill").astype(int)
    excel[KOUDOU_ID] = excel[KOUDOU_ID].fillna(method="ffill").astype(int)
    excel[KOUTEI] = excel[KOUTEI].fillna(method="ffill")
    excel[SAGYO] = excel[SAGYO].fillna(method="ffill")
    sub_IDs = excel[colmuns]
    IDs = IDs.append(sub_IDs)
IDs = IDs.drop_duplicates().sort_values([KOUTEI_ID,SAGYO_ID,KOUDOU_ID]).reset_index(drop=True)

#エクセル作成(技能分析表本体)
wb = xlApp.Workbooks.Add()
ws = wb.sheets[0]
ws.Name = "比較技能分析表"
ws.Cells.Font.Name = "メイリオ"
ws.Cells.Font.size = 10
#折り返して表示
ws.Cells.WrapText = True
#エクセルが見える状態で編集
xlApp.Visible = True
#警告を表示しない(Trueでも問題ない)
xlApp.DisplayAlerts = False

#header作成
header = [
    "工程",
    "作業",
    "具体的な行動\n共通人数",
    "行動のポイント\n共通人数",
    "行動の背景\n共通人数",
    "比較"
]
#headerを記入
ws.Range(ws.Cells(2,2),ws.Cells(2,7)).Value = header
ws.Range(ws.Cells(2,2),ws.Cells(2,7)).HorizontalAlignment = 3
ws.Range(ws.Cells(2,2),ws.Cells(2,7)).Interior.ColorIndex = 15
ws.Range("D2:F2").ColumnWidth = 12
ws.Range("G2").ColumnWidth = 24
for i, sheet_name in enumerate(sheet_names):
    i += 1 #excelの列インデックスに使うため
    ws.Cells(2,i*3+INDEX_PLUS).Value = sheet_name
    target_range = ws.Range(ws.Cells(2,i*3+INDEX_PLUS),ws.Cells(2,i*3+INDEX_PLUS+2))
    target_range.MergeCells = True
    target_range.HorizontalAlignment = 3
    target_range.ColumnWidth = 30
    target_range.Interior.ColorIndex = 15

#技能分析表作成
koutei_ID = 1
sagyo_ID = 1
koutei_st = 3
sagyo_st = 3
next_row = 3

#はじめの作業、工程列を記入
ws.Cells(koutei_st,2).Value = IDs[KOUTEI][0]
ws.Cells(sagyo_st,3).value = IDs[SAGYO][0]

for row in range(len(IDs)):
    #koutei_IDが変わった時に工程列をセル結合し、工程を記入
    if koutei_ID != IDs[KOUTEI_ID][row]:
        #工程列をセル結合
        target_range = ws.Range(
            ws.Cells(koutei_st,2),
            ws.Cells(next_row-1,2)
        )
        target_range.MergeCells = True
        target_range.HorizontalAlignment = 3
        #koutei_st更新　next_rowと同じ(next_row=次記入する行)
        koutei_st = next_row
        #次の工程記入
        ws.Cells(koutei_st,2).Value = IDs[KOUTEI][row]
        
    #sagyo_idが変わった時に作業列セル結合し、next_rowを更新
    if sagyo_ID != IDs[SAGYO_ID][row] or koutei_ID != IDs[KOUTEI_ID][row]:
        #作業列セル結合
        target_range = ws.Range(
            ws.Cells(sagyo_st,3),
            ws.Cells(next_row-1,3)
        )
        target_range.MergeCells = True
        target_range.HorizontalAlignment = 3
        #比較列セル結合
        target_range = ws.Range(
            ws.Cells(sagyo_st,7),
            ws.Cells(next_row-1,7)
        )
        target_range.MergeCells = True
        target_range.HorizontalAlignment = 3
        #sagyo_st更新
        sagyo_st = next_row
        #次の作業記入
        ws.Cells(sagyo_st,3).value = IDs[SAGYO][row]
        
    #koutei_ID更新
    koutei_ID = IDs[KOUTEI_ID][row]
    #sagyo_ID更新
    sagyo_ID = IDs[SAGYO_ID][row]

    for sub_cnt, sheet_name in enumerate(sheet_names):
        sub_cnt += 1 #excelの列インデックスに使うため
        excel = excels[sheet_name]
        #koutei_ID,sagyou_ID,Koudou_IDが全て一致するものを抽出
        skill_table = excel[
            (excel[KOUTEI_ID] == IDs[KOUTEI_ID][row]) &
            (excel[SAGYO_ID] == IDs[SAGYO_ID][row]) &
            (excel[KOUDOU_ID] == IDs[KOUDOU_ID][row])
        ]
        
        #動画があれば、リンクを挿入
        df_video = skill_table.filter(like = "video", axis = 1).reset_index(drop=True)
        df_video = df_video.dropna(how="all", axis=1)
        if(not(df_video.empty)):
            
            for video_cnt,video_index in enumerate(df_video):
                video_path = os.getcwd() + "/videos/" + df_video[video_index][0]
                #ディレクトリにファイルが存在する場合はリンクを挿入
                if os.path.exists(video_path):
                    ws.Cells(next_row, sub_cnt*3+5+video_cnt).Value = df_video[video_index][0]
                    ws.Cells(next_row, sub_cnt*3+5+video_cnt).HorizontalAlignment = 3
                    ws.Hyperlinks.Add(
                        Anchor = ws.Cells(next_row, sub_cnt*3+INDEX_PLUS+video_cnt), 
                        Address = video_path
                    )
                else:
                    print(video_path, "がありません。")
        #画像があれば挿入
        df_image = skill_table.filter(like = "image", axis = 1)
        df_image = df_image.dropna(how="all", axis=1).reset_index(drop=True)
        if(not(df_image.empty)):
            
            image_row = next_row + 1#画像を挿入する行数
            #画像行の行高さを調整 img_size*33
            ws.Rows(image_row).RowHeight = img_size * 33 
            for image_cnt,image_index in enumerate(df_image):
                image_path = os.getcwd() + "/images/" + df_image[image_index][0]
                #画像3枚ごとに次の行に挿入
                if image_cnt / 3 == 1:
                    image_row += 1
                    #画像行の行高さを調整 img_size*33
                    ws.Rows(image_row).RowHeight = img_size * 33
                #ディレクトリにファイルが存在する場合は画像を挿入
                if os.path.exists(image_path):
                    image = ws.Shapes.AddPicture(
                        Filename = os.getcwd() + "/images/" + df_image[image_index][0],
                        LinkToFile = False,
                        SaveWithDocument = True,
                        Left = ws.Cells(image_row, sub_cnt*3+INDEX_PLUS+image_cnt).Left,
                        Top = ws.Cells(image_row, sub_cnt*3+INDEX_PLUS+image_cnt).Top,
                        Width = -1,
                        Height= -1
                    )
                    #画像番号入力
                    ws.Cells(image_row, sub_cnt*3+INDEX_PLUS+image_cnt).value = image_cnt + 1
                    ws.Cells(image_row, sub_cnt*3+INDEX_PLUS+image_cnt).HorizontalAlignment = 2
                    ws.Cells(image_row, sub_cnt*3+INDEX_PLUS+image_cnt).VerticalAlignment = 1
                    
                    #画像サイズ調整
                    image.Height = 28.34646 * img_size
                    if image.Width > 28.34646 * 6:#img_sizeで指定した幅(cm)のアスペクト比計算で幅が6cmを超える場合は7cmに設定
                        image.Width = 28.34646 * 6
                    
                    #画像をセルの中央へ配置
                    image.Left = ws.Cells(image_row,sub_cnt*3+INDEX_PLUS+image_cnt).Left \
                        + (ws.Cells(image_row,sub_cnt*3+INDEX_PLUS+image_cnt).Width - image.Width) / 2
                    image.Top = ws.Cells(image_row,sub_cnt*3+INDEX_PLUS+image_cnt).Top\
                        + (ws.Cells(image_row,sub_cnt*3+INDEX_PLUS+image_cnt).Height - image.Height) / 2
                else:
                    print(image_path, "がありません。")
    
    #画像と動画を入れ終えてから
    #koudou_idが1の時はミニヘッダーを挿入       
    if IDs[KOUDOU_ID][row] == 1:
        #miniheader記入
        minihead_row = ws.UsedRange.Rows.Count + 2
        miniheader = ["具体的な行動", "行動のポイント", "行動の背景"]
        #被験者分
        for i in range(len(sheet_names)):
            i += 1 #excelの列インデックスに使うため
            target_range = ws.Range(
                ws.Cells(
                    minihead_row,
                    i*3+INDEX_PLUS
                ),
                ws.Cells(
                    minihead_row,
                    i*3+INDEX_PLUS+2)
            )
            target_range.Value = miniheader
            target_range.HorizontalAlignment = 3
            target_range.Interior.ColorIndex = 15
                
    #技能分析表の内容を記入
    skill_table_row = ws.UsedRange.Rows.Count + 2
    #一人ずつ内容を埋めていく
    for sub_cnt, sheet_name in enumerate(sheet_names):
        sub_cnt += 1 #excelの列インデックスに使うため
        excel = excels[sheet_name]
        #koutei_ID,sagyou_ID,Koudou_IDが全て一致するものを抽出
        skill_table = excel[
            (excel[KOUTEI_ID] == IDs[KOUTEI_ID][row]) &
            (excel[SAGYO_ID] == IDs[SAGYO_ID][row]) &
            (excel[KOUDOU_ID] == IDs[KOUDOU_ID][row])
        ]
        #内容記載に必要部分のみ抽出
        skill_table = skill_table.loc[:, ["具体的な行動の仕方", "行動のポイント", "ポイントの背景"]].fillna("")
        #行動の行数
        num_koudou = skill_table.shape[0]
        #内容を埋めていく
        if num_koudou == 1:
            ws.Range(
                ws.Cells(
                    skill_table_row,
                    sub_cnt*3+INDEX_PLUS
                ),
                ws.Cells(
                    skill_table_row,
                    sub_cnt*3+INDEX_PLUS+2)
            ).value = skill_table.values.tolist()
        else:
            ws.Range(
                ws.Cells(
                    skill_table_row,
                    sub_cnt*3+INDEX_PLUS
                ),
                ws.Cells(
                    skill_table_row+num_koudou,
                    sub_cnt*3+INDEX_PLUS+2)
            ).value = skill_table.values.tolist()
    next_row = ws.UsedRange.Rows.Count + 2
    
#最後のセル結合
#工程列をセル結合
target_range = ws.Range(
    ws.Cells(koutei_st,2),
    ws.Cells(next_row-1,2)
)
target_range.MergeCells = True
target_range.HorizontalAlignment = 3
#作業列セル結合
target_range = ws.Range(
    ws.Cells(sagyo_st,3),
    ws.Cells(next_row-1,3)
)
target_range.MergeCells = True
target_range.HorizontalAlignment = 3
#比較列セル結合
target_range = ws.Range(
    ws.Cells(sagyo_st,7),
    ws.Cells(next_row-1,7)
)
target_range.MergeCells = True
target_range.HorizontalAlignment = 3

#体裁を整える
#罫線
ws.UsedRange.Borders.LineStyle = 1
for i in range(1,len(sheet_names)):
    ws.Range(
        ws.Cells(2,i*3+7),
        ws.Cells(ws.UsedRange.Rows.Count+2,i*3+7)
        ).Borders(2).LineStyle = -4119


#保存用ディレクトリ作成
if not os.path.exists("output"):
    os.mkdir("output")
#エクセルを保存
save_file = save_name + "_" + str(datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')) + ".xlsx"
wb.SaveAs(f"{os.getcwd()}\\output\\{save_file}", FileFormat = 51)
wb.Close()
xlApp.Quit()