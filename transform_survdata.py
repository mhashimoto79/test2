import os # ファイル名関係
import pandas as pd
import re
import csv
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

def horizontal_to_vertical(df, target_col, id_col):
    """
    機能：
    ・横持ちデータを縦持ちデータに変換
    引数：
    ・df：データフレーム
    ・id_col：IDカラム名
    戻り値：
    ・df_vertical：縦持ちデータ
    メモ：
    ・
    残件：
    ・
    """
    df_vertical = df[target_col]
    df_vertical = df_vertical.stack().reset_index()
    df_vertical = df_vertical.rename(columns={'level_0':"ID",'level_1':'Q_ID',0:'A'})
    df_vertical = pd.merge(df[id_col],df_vertical, left_index=True, right_on='ID')
    df_vertical = df_vertical.drop(columns=['ID'])
    return df_vertical

def edit_Atocos(qrawFile, qlayoutFile, qeditDir):
    """
    機能：
    ・Atocos用データ加工
    引数：
    ・qrawFile：アンケートパス(絶対パス)(ObjectType:str)
    ・qlayoutFile：アンケートフォーマットファイルパス(絶対パス)(ObjectType:str)
    ・qeditDir：加工データ出力フォルダパス(ObjectType:str)
    メモ：
    残件：
    ・メッセージ
    """
    #変数定義
    id_col = 'ANSWER_ID'
    qid_col = '質問ID'
    qtype_col = '質問タイプ'
    atype_col = '回答タイプ'
    label_col  = 'カラムID'
    iname_col = 'アイテム名'
    choice_col = '質問文/選択肢'
    choicenum_col = '選択肢番号'
    qtype_fa = 'FA'
    qtype_da = 'DATE'
    qtype_na = 'NA'
    qtype_sa= 'SA'
    qtype_ma= 'MA'
    qtype_mtx = 'MATRIX' #複数のシングル/マルチ
    qtype_mti = 'MULTIPLE_INPUT' #複数のフリー記載
    atype_fa = 'FA' #フリー
    atype_sa = 'SA' #シングル
    atype_ma = 'MA' #マルチ
    atype_na = 'NA' #番号
    qtype_da = 'NA' #日付

    #rawファイル編集：Atocosの列名が不足するケースがあるため
    with open(qrawFile,mode='r',encoding='shift_jis') as rf:
        reader = csv.reader(rf)
        rows = list(reader)
        if rows[0][2] == 'OS':
            rows[0].insert(1,'ANSWER_DATE')
            rows[0].insert(2,'SURVEY_VERSION')
        rf.close()

    with open('_tmp.csv',mode='w',encoding='shift_jis') as wf:
        writer = csv.writer(wf)
        writer.writerows(rows)
        wf.close()

    #ファイル読み込み
    df_layout = pd.read_csv(qlayoutFile, encoding="shift-jis", keep_default_na=False, na_values=[''])
    df_raw = pd.read_csv('_tmp.csv', encoding="shift-jis", keep_default_na=False, na_values=['']) 

    #tmpのrawファイル削除
    os.remove('_tmp.csv')

    #データ加工
    #layoutデータ加工
    df_layout[qid_col].ffill(inplace=True)
    df_layout[qtype_col].ffill(inplace=True)
    df_layout[atype_col].ffill(inplace=True)
    #rawデータ加工
    for i in range(0, len(df_raw.columns)):
        col_name = df_raw.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        if not len(df_tmp) ==0:
            if re.match(r'^(QS).+$', df_tmp[label_col].values[0]):
                if df_tmp[iname_col].isna().values[0] or re.match(r'.+TEXT$', df_tmp[label_col].values[0]) or re.match(r'.+T\d+$', df_tmp[label_col].values[0]):
                    tmp_colname = re.sub(r'^QS[0-9]+', df_tmp[qid_col].values[0], df_tmp[label_col].values[0])
                    df_raw = df_raw.rename(columns={df_tmp[label_col].values[0]:tmp_colname})      
                else:
                    df_raw = df_raw.rename(columns={df_tmp[label_col].values[0]:df_tmp[iname_col].values[0]})
    #layoutデータ加工：2回目
    for i in range(0, len(df_layout)):
            if pd.notnull(df_layout.iloc[i][label_col]):
                if re.match(r'^QS[0-9]+', df_layout.iloc[i][label_col]):
                    df_layout.loc[i,label_col] = re.sub(r'^QS[0-9]+', df_layout.iloc[i][qid_col], df_layout.iloc[i][label_col])

    # labelデータ作成
    df_label = df_raw
    for i in range(0, len(df_raw.columns)):
        col_name = df_raw.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        if df_tmp[atype_col].values[0] == atype_ma:#回答タイプがマルチアンサーの場合
            label = df_tmp[choice_col].values[0]
            df_label[col_name] = df_raw[col_name].replace(1, label).replace(0, None)
        elif df_tmp[atype_col].values[0] == atype_sa:#回答タイプがシングルアンサーの場合
            qid = df_tmp[qid_col].values[0]
            df_tmp = df_layout[df_layout[qid_col] == qid]
            df_tmp[label_col].ffill(inplace=True)
            df_tmp = df_tmp[df_tmp[label_col] == col_name]
            df_tmp[choicenum_col] = df_tmp[choicenum_col].astype(pd.Int64Dtype(), errors='ignore')
            df_tmp_dict = df_tmp.set_index(choicenum_col)[choice_col].to_dict()
            df_label[col_name] = df_raw[col_name].map(df_tmp_dict)
    
    # カラム名抽出
    #質問タイプ：シングルアンサー
    col_list_sa = df_layout[df_layout[qtype_col].isin([qtype_sa])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_sa = [x for x in col_list_sa if re.match(r'^(Q|SC)\d+$', x)]
    col_list_qtype_sa_txt = [x for x in col_list_sa if re.match(r'^(Q|SC).+TEXT$', x)]
    col_list_kihon_sa = [x for x in col_list_sa if x not in col_list_qtype_sa and x not in col_list_qtype_sa_txt]
    #質問タイプ：マルチアンサー
    col_list_ma = df_layout[df_layout[qtype_col].isin([qtype_ma])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_ma = [x for x in col_list_ma if re.match(r'^(Q|SC).+\d+$', x)]
    col_list_qtype_ma_txt = [x for x in col_list_ma if re.match(r'^(Q|SC).+TEXT$', x)]
    col_list_kihon_ma = [x for x in col_list_ma if x not in col_list_qtype_ma and x not in col_list_qtype_ma_txt]
    #質問タイプ：MTI
    col_list_mti = df_layout[df_layout[qtype_col].isin([qtype_mti])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_mti = [x for x in col_list_mti if re.match(r'^(Q|SC).+\d+$', x)]
    #質問タイプ：マトリックス
    col_list_mtx = df_layout[df_layout[qtype_col].isin([qtype_mtx])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_mtx = [x for x in col_list_mtx if re.match(r'^(Q|SC).+\d+$', x)]
    col_list_kihon_mtx = [x for x in col_list_mtx if x not in col_list_qtype_mtx]

    #データ作成：縦持ちデータ
    #質問タイプ：シングルアンサー
    df_label_sa = horizontal_to_vertical(df_label,col_list_qtype_sa,id_col)
    df_label_sa_txt = horizontal_to_vertical(df_label,col_list_qtype_sa_txt,id_col)
    #df_label_sa_kihon = horizontal_to_vertical(df_label,col_list_kihon_sa,id_col)
    #df_label_sa = pd.concat([df_label_sa,df_label_sa_kihon],axis=0, ignore_index=True)
    df_label_sa.insert(2,'Q',df_label_sa['Q_ID'])
    df_label_sa['Q']= df_label_sa['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_sa_txt['Q_ID'] = df_label_sa_txt['Q_ID'].str.replace(r'_\d+_TEXT', '',regex=True)
    df_label_sa['TEXT'] = pd.merge(df_label_sa, df_label_sa_txt, on=[id_col, 'Q_ID'], how='left')['A_y']
    #質問タイプ：マルチアンサー
    df_label_ma = horizontal_to_vertical(df_label,col_list_qtype_ma,id_col)
    df_label_ma_txt = horizontal_to_vertical(df_label,col_list_qtype_ma_txt,id_col)
    df_label_ma_kihon = horizontal_to_vertical(df_label,col_list_kihon_ma,id_col)
    df_label_ma = pd.concat([df_label_ma,df_label_ma_kihon],axis=0, ignore_index=True)
    df_label_ma['Q_ID'] = df_label_ma['Q_ID'].str.replace(r'_\d+', '',regex=True)
    df_label_ma.insert(2,'Q',df_label_ma['Q_ID'])
    df_label_ma['Q']= df_label_ma['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_ma_txt['Q_ID'] = df_label_ma_txt['Q_ID'].str.replace(r'_\d+_TEXT', '',regex=True)
    df_label_ma['TEXT'] = pd.merge(df_label_ma, df_label_ma_txt, on=[id_col, 'Q_ID'], how='left')['A_y']
    #質問タイプ：MTI
    df_label_mti = horizontal_to_vertical(df_label,col_list_qtype_mti,id_col)
    df_label_mti['Q_ID'] = df_label_mti['Q_ID'].str.replace(r'_T\d+', '',regex=True)
    df_label_mti.insert(2,'Q',df_label_mti['Q_ID'])
    df_label_mti['Q']= df_label_mti['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    #結合
    df_label_q = pd.concat([df_label_sa, df_label_ma, df_label_mti], axis=0, ignore_index=True).sort_values([id_col, 'Q_ID'])
    #質問タイプ：マトリックス
    df_label_mtx = horizontal_to_vertical(df_label,col_list_qtype_mtx,id_col)
    df_label_mtx_kihon = horizontal_to_vertical(df_label,col_list_kihon_mtx,id_col)
    df_label_mtx = pd.concat([df_label_mtx,df_label_mtx_kihon],axis=0, ignore_index=True)
    df_label_mtx.insert(2,'Q_ID_2',df_label_mtx['Q_ID'].str.replace(r'_T\d+', '',regex=True))
    df_label_mtx.insert(3,'Q_ID_3',df_label_mtx['Q_ID'].str.replace(r'(Q|SC)\d+_', '',regex=True))
    df_label_mtx.insert(4,'Q',df_label_mtx['Q_ID_2'])
    df_label_mtx['Q']= df_label_mtx['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_mtx.insert(5,'Q_2',df_label_mtx['Q_ID'])
    df_label_mtx['Q_2']= df_label_mtx['Q_2'].replace(df_layout[label_col].to_list(), df_layout[choice_col])
    df_label_mtx = df_label_mtx.sort_values([id_col, 'Q_ID'])

    #データ作成：横持ちデータ
    df_label_kihon = df_label.drop(col_list_qtype_ma+col_list_qtype_ma_txt+col_list_kihon_ma+col_list_qtype_mtx+col_list_kihon_mtx+col_list_qtype_mti,axis=1)
    #カラム名変更：英語→日本語
    for i in range(0, len(df_label_kihon.columns)):
        col_name = df_label_kihon.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        df_label_kihon = df_label_kihon.rename(columns={col_name:df_tmp[choice_col].values[0]})

    #出力
    qrawFileName = os.path.splitext(os.path.basename(qrawFile))[0]
    qlayoutFileName = os.path.splitext(os.path.basename(qlayoutFile))[0]
    df_label_kihon.to_csv(qeditDir+'/'+qrawFileName+'_out1.csv', index = False)
    df_label_q.to_csv(qeditDir+'/'+qrawFileName+'_out2.csv', index = False)
    df_label_mtx.to_csv(qeditDir+'/'+qrawFileName+'_out3.csv', index = False)
    df_label_kihon.to_excel(qeditDir+'/'+qrawFileName+'_out1.xlsx', index = False)
    df_label_q.to_excel(qeditDir+'/'+qrawFileName+'_out2.xlsx', index = False)
    df_label_mtx.to_excel(qeditDir+'/'+qrawFileName+'_out3.xlsx', index = False)
    df_layout.to_excel(qeditDir+'/'+qlayoutFileName+'_edit.xlsx', index = False)
    df_label.to_excel(qeditDir+'/'+qrawFileName+'_edit.xlsx', index = False)

def edit_MApps(qrawFile, qlayoutFile, qeditDir):
    """
    機能：
    ・Atocos用データ加工
    引数：
    ・qrawFile：アンケートパス(絶対パス)(ObjectType:str)
    ・qlayoutFile：アンケートフォーマットファイルパス(絶対パス)(ObjectType:str)
    ・qeditDir：加工データ出力フォルダパス(ObjectType:str)
    メモ：
    残件：
    ・メッセージ
    """
    #変数定義
    id_col = 'ANSWER_ID'
    qid_col = '質問ID'
    qtype_col = '質問タイプ'
    atype_col = '回答タイプ'
    label_col  = 'カラムID'
    iname_col = 'アイテム名'
    choice_col = '質問文/選択肢'
    choicenum_col = '選択肢番号'
    qtype_fa = 'FA'
    qtype_da = 'DATE'
    qtype_na = 'NA'
    qtype_sa= 'SA'
    qtype_ma= 'MA'
    qtype_mtx = 'MATRIX' #複数のシングル/マルチ
    qtype_mti = 'MULTIPLE_INPUT' #複数のフリー記載
    atype_fa = 'FA' #フリー
    atype_sa = 'SA' #シングル
    atype_ma = 'MA' #マルチ
    atype_na = 'NA' #番号
    qtype_da = 'NA' #日付

    #rawファイル編集：Atocosの列名が不足するケースがあるため
    with open(qrawFile,mode='r',encoding='shift_jis') as rf:
        reader = csv.reader(rf)
        rows = list(reader)
        if rows[0][2] == 'OS':
            rows[0].insert(1,'ANSWER_DATE')
            rows[0].insert(2,'SURVEY_VERSION')
        rf.close()

    with open('_tmp.csv',mode='w',encoding='shift_jis') as wf:
        writer = csv.writer(wf)
        writer.writerows(rows)
        wf.close()

    #ファイル読み込み
    df_layout = pd.read_csv(qlayoutFile, encoding="shift-jis", keep_default_na=False, na_values=[''])
    df_raw = pd.read_csv('_tmp.csv', encoding="shift-jis", keep_default_na=False, na_values=['']) 

    #tmpのrawファイル削除
    os.remove('_tmp.csv')

    #データ加工
    #layoutデータ加工
    df_layout[qid_col].ffill(inplace=True)
    df_layout[qtype_col].ffill(inplace=True)
    df_layout[atype_col].ffill(inplace=True)
    #rawデータ加工
    for i in range(0, len(df_raw.columns)):
        col_name = df_raw.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        if not len(df_tmp) ==0:
            if re.match(r'^(QS).+$', df_tmp[label_col].values[0]):
                if df_tmp[iname_col].isna().values[0] or re.match(r'.+TEXT$', df_tmp[label_col].values[0]) or re.match(r'.+T\d+$', df_tmp[label_col].values[0]):
                    tmp_colname = re.sub(r'^QS[0-9]+', df_tmp[qid_col].values[0], df_tmp[label_col].values[0])
                    df_raw = df_raw.rename(columns={df_tmp[label_col].values[0]:tmp_colname})      
                else:
                    df_raw = df_raw.rename(columns={df_tmp[label_col].values[0]:df_tmp[iname_col].values[0]})
    #layoutデータ加工：2回目
    for i in range(0, len(df_layout)):
            if pd.notnull(df_layout.iloc[i][label_col]):
                if re.match(r'^QS[0-9]+', df_layout.iloc[i][label_col]):
                    df_layout.loc[i,label_col] = re.sub(r'^QS[0-9]+', df_layout.iloc[i][qid_col], df_layout.iloc[i][label_col])

    # labelデータ作成
    df_label = df_raw
    for i in range(0, len(df_raw.columns)):
        col_name = df_raw.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        if df_tmp[atype_col].values[0] == atype_ma:#回答タイプがマルチアンサーの場合
            label = df_tmp[choice_col].values[0]
            df_label[col_name] = df_raw[col_name].replace(1, label).replace(0, None)
        elif df_tmp[atype_col].values[0] == atype_sa:#回答タイプがシングルアンサーの場合
            qid = df_tmp[qid_col].values[0]
            df_tmp = df_layout[df_layout[qid_col] == qid]
            df_tmp[label_col].ffill(inplace=True)
            df_tmp = df_tmp[df_tmp[label_col] == col_name]
            df_tmp[choicenum_col] = df_tmp[choicenum_col].astype(pd.Int64Dtype(), errors='ignore')
            df_tmp_dict = df_tmp.set_index(choicenum_col)[choice_col].to_dict()
            df_label[col_name] = df_raw[col_name].map(df_tmp_dict)
    
    # カラム名抽出
    #質問タイプ：シングルアンサー
    col_list_sa = df_layout[df_layout[qtype_col].isin([qtype_sa])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_sa = [x for x in col_list_sa if re.match(r'^(Q|SC)\d+$', x)]
    col_list_qtype_sa_txt = [x for x in col_list_sa if re.match(r'^(Q|SC).+TEXT$', x)]
    col_list_kihon_sa = [x for x in col_list_sa if x not in col_list_qtype_sa and x not in col_list_qtype_sa_txt]
    #質問タイプ：マルチアンサー
    col_list_ma = df_layout[df_layout[qtype_col].isin([qtype_ma])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_ma = [x for x in col_list_ma if re.match(r'^(Q|SC).+\d+$', x)]
    col_list_qtype_ma_txt = [x for x in col_list_ma if re.match(r'^(Q|SC).+TEXT$', x)]
    col_list_kihon_ma = [x for x in col_list_ma if x not in col_list_qtype_ma and x not in col_list_qtype_ma_txt]
    #質問タイプ：MTI
    col_list_mti = df_layout[df_layout[qtype_col].isin([qtype_mti])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_mti = [x for x in col_list_mti if re.match(r'^(Q|SC).+\d+$', x)]
    #質問タイプ：マトリックス
    col_list_mtx = df_layout[df_layout[qtype_col].isin([qtype_mtx])][label_col].drop_duplicates().dropna().tolist()
    col_list_qtype_mtx = [x for x in col_list_mtx if re.match(r'^(Q|SC).+\d+$', x)]
    col_list_kihon_mtx = [x for x in col_list_mtx if x not in col_list_qtype_mtx]

    #データ作成：縦持ちデータ
    #質問タイプ：シングルアンサー
    df_label_sa = horizontal_to_vertical(df_label,col_list_qtype_sa,id_col)
    df_label_sa_txt = horizontal_to_vertical(df_label,col_list_qtype_sa_txt,id_col)
    #df_label_sa_kihon = horizontal_to_vertical(df_label,col_list_kihon_sa,id_col)
    #df_label_sa = pd.concat([df_label_sa,df_label_sa_kihon],axis=0, ignore_index=True)
    df_label_sa.insert(2,'Q',df_label_sa['Q_ID'])
    df_label_sa['Q']= df_label_sa['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_sa_txt['Q_ID'] = df_label_sa_txt['Q_ID'].str.replace(r'_\d+_TEXT', '',regex=True)
    df_label_sa['TEXT'] = pd.merge(df_label_sa, df_label_sa_txt, on=[id_col, 'Q_ID'], how='left')['A_y']
    #質問タイプ：マルチアンサー
    df_label_ma = horizontal_to_vertical(df_label,col_list_qtype_ma,id_col)
    df_label_ma_txt = horizontal_to_vertical(df_label,col_list_qtype_ma_txt,id_col)
    df_label_ma_kihon = horizontal_to_vertical(df_label,col_list_kihon_ma,id_col)
    df_label_ma = pd.concat([df_label_ma,df_label_ma_kihon],axis=0, ignore_index=True)
    df_label_ma['Q_ID'] = df_label_ma['Q_ID'].str.replace(r'_\d+', '',regex=True)
    df_label_ma.insert(2,'Q',df_label_ma['Q_ID'])
    df_label_ma['Q']= df_label_ma['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_ma_txt['Q_ID'] = df_label_ma_txt['Q_ID'].str.replace(r'_\d+_TEXT', '',regex=True)
    df_label_ma['TEXT'] = pd.merge(df_label_ma, df_label_ma_txt, on=[id_col, 'Q_ID'], how='left')['A_y']
    #質問タイプ：MTI
    df_label_mti = horizontal_to_vertical(df_label,col_list_qtype_mti,id_col)
    df_label_mti['Q_ID'] = df_label_mti['Q_ID'].str.replace(r'_T\d+', '',regex=True)
    df_label_mti.insert(2,'Q',df_label_mti['Q_ID'])
    df_label_mti['Q']= df_label_mti['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    #結合
    df_label_q = pd.concat([df_label_sa, df_label_ma, df_label_mti], axis=0, ignore_index=True).sort_values([id_col, 'Q_ID'])
    #質問タイプ：マトリックス
    df_label_mtx = horizontal_to_vertical(df_label,col_list_qtype_mtx,id_col)
    df_label_mtx_kihon = horizontal_to_vertical(df_label,col_list_kihon_mtx,id_col)
    df_label_mtx = pd.concat([df_label_mtx,df_label_mtx_kihon],axis=0, ignore_index=True)
    df_label_mtx.insert(2,'Q_ID_2',df_label_mtx['Q_ID'].str.replace(r'_T\d+', '',regex=True))
    df_label_mtx.insert(3,'Q_ID_3',df_label_mtx['Q_ID'].str.replace(r'(Q|SC)\d+_', '',regex=True))
    df_label_mtx.insert(4,'Q',df_label_mtx['Q_ID_2'])
    df_label_mtx['Q']= df_label_mtx['Q'].replace(df_layout[qid_col].iloc[::-1].to_list(), df_layout[choice_col].iloc[::-1])
    df_label_mtx.insert(5,'Q_2',df_label_mtx['Q_ID'])
    df_label_mtx['Q_2']= df_label_mtx['Q_2'].replace(df_layout[label_col].to_list(), df_layout[choice_col])
    df_label_mtx = df_label_mtx.sort_values([id_col, 'Q_ID'])

    #データ作成：横持ちデータ
    df_label_kihon = df_label.drop(col_list_qtype_ma+col_list_qtype_ma_txt+col_list_kihon_ma+col_list_qtype_mtx+col_list_kihon_mtx+col_list_qtype_mti,axis=1)
    #カラム名変更：英語→日本語
    for i in range(0, len(df_label_kihon.columns)):
        col_name = df_label_kihon.columns[i]
        df_tmp = df_layout[df_layout[label_col] == col_name]
        df_label_kihon = df_label_kihon.rename(columns={col_name:df_tmp[choice_col].values[0]})

    #出力
    qrawFileName = os.path.splitext(os.path.basename(qrawFile))[0]
    qlayoutFileName = os.path.splitext(os.path.basename(qlayoutFile))[0]
    df_label_kihon.to_csv(qeditDir+'/'+qrawFileName+'_out1.csv', index = False)
    df_label_q.to_csv(qeditDir+'/'+qrawFileName+'_out2.csv', index = False)
    df_label_mtx.to_csv(qeditDir+'/'+qrawFileName+'_out3.csv', index = False)
    df_label_kihon.to_excel(qeditDir+'/'+qrawFileName+'_out1.xlsx', index = False)
    df_label_q.to_excel(qeditDir+'/'+qrawFileName+'_out2.xlsx', index = False)
    df_label_mtx.to_excel(qeditDir+'/'+qrawFileName+'_out3.xlsx', index = False)
    df_layout.to_excel(qeditDir+'/'+qlayoutFileName+'_edit.xlsx', index = False)
    df_label.to_excel(qeditDir+'/'+qrawFileName+'_edit.xlsx', index = False)

def edit_survey(qrawFile, qlayoutFile, qeditDir, category):
    """
    機能：
    ・データ加工のメイン
    引数：
    ・qrawFile：アンケートパス(絶対パス)(ObjectType:str)
    ・qlayoutFile：アンケートフォーマットファイルパス(絶対パス)(ObjectType:str)
    ・qeditDir：加工データ出力フォルダパス(ObjectType:str)
    ・category：アンケートフォーマットのカテゴリ(ObjectType:str)
    メモ：
    ・pythonでxlsxファイルを操作するライブラリは「xlwings」「openpyxl」がある
    　比較サイト：https://posipochi.com/2021/05/29/python-opnepyxl-xlwings/
    ・openpyxl　は　軽量、Excel未インストールでも操作可能
    ・xlwings　は　多機能、Excel未インストールだと操作不可
    残件：
    ・メッセージ
    """
    #ファイル読み込み
    if category == 'Atcos':
        edit_Atocos(qrawFile, qlayoutFile, qeditDir)
    elif category == 'MApps':
        #edit_Atocos(qrawFile, qlayoutFile, qeditDir)
        tk.messagebox.showinfo('transform_survdata：確認','機能が実装されていません。')
    elif category == 'LINE':
        #edit_Atocos(qrawFile, qlayoutFile, qeditDir)
        tk.messagebox.showinfo('transform_survdata：確認','機能が実装されていません。')
    else:
        raise SyntaxError

    # シート毎に処理を実施
    # Close処理
    tk.messagebox.showinfo('transform_survdata：確認','処理が正常に完了しました。')

# 実行ボタン押下時に処理を実行する関数
def run_func(file_set_frame_qraw, file_set_frame_qlayout, dir_set_frame_qedit, pulldown_frame):
    try:
        qrawFile = file_set_frame_qraw.edit_box.get()
        qlayoutFile = file_set_frame_qlayout.edit_box.get()
        qeditDir = dir_set_frame_qedit.edit_box.get()
        category = pulldown_frame.combobox.get()
        edit_survey(qrawFile, qlayoutFile, qeditDir, category)
    except Exception as e:
        tk.messagebox.showerror('transform_survdata：エラー','エラーが発生しました。ファイル or フォルダ の指定が正しいか確認してください')
        #デバッグ用
        #tk.messagebox.showerror('transform_survdata：エラー',e)
# 実行エリアのフレームを作成して返却する関数
def run_funct_frame(parent_frame, file_set_frame_qraw, file_set_frame_qlayout, dir_set_frame_qedit, pulldown_frame, label_text):
    run_frame = ttk.Frame(parent_frame)
    run_button = tk.Button(run_frame, text = label_text, width = 10\
    , command = lambda:run_func(file_set_frame_qraw, file_set_frame_qlayout, dir_set_frame_qedit, pulldown_frame))
    run_button.pack(side = tk.LEFT)
    return run_frame
#  [FILE]ボタン押下時に呼び出し。選択したファイルのパスをテキストボックスに設定する。
def open_file_command(edit_box, file_type_list):
    iDir = os.path.abspath(os.path.dirname(__file__))
    file_path = filedialog.askopenfilename(filetypes = file_type_list,initialdir = iDir)
    edit_box.delete(0, tk.END)
    edit_box.insert(tk.END, file_path)
#  [FOLDER]ボタン押下時に呼び出し。選択したファイルのパスをテキストボックスに設定する。
def open_dir_command(edit_box):
    iDir = os.path.abspath(os.path.dirname(__file__))
    dir_path = filedialog.askdirectory(initialdir = iDir)
    edit_box.delete(0, tk.END)
    edit_box.insert(tk.END, dir_path)
#プルダウンエリアのフレームを作成して返却する関数
def set_pulldown_frame(parent_frame, label_text, pulldown_list):
    pulldown_frame = ttk.Frame(parent_frame)
    tk.Label(pulldown_frame, text = label_text).pack(side = tk.LEFT)
    # プルダウンの作成と配置
    pulldown_frame.combobox = ttk.Combobox(pulldown_frame, state = 'readonly', width = 50)
    pulldown_frame.combobox['values'] = pulldown_list
    pulldown_frame.combobox.pack(side = tk.LEFT)
    pulldown_frame.combobox.set(pulldown_list[0])
    return pulldown_frame
# フォルダ設定エリアのフレームを作成して返却する関数
def set_dir_frame(parent_frame, label_text):
    dir_frame = ttk.Frame(parent_frame)
    tk.Label(dir_frame, text = label_text).pack(side = tk.LEFT)
    # テキストボックスの作成と配置
    dir_frame.edit_box = tk.Entry(dir_frame, width = 50)
    dir_frame.edit_box.pack(side = tk.LEFT)
    # ボタンの作成と配置
    dir_button = tk.Button(dir_frame, text = 'FOLDER', width = 5\
        , command = lambda:open_dir_command(dir_frame.edit_box))
    dir_button.pack(side = tk.LEFT)
    return dir_frame
# ファイル設定エリアのフレームを作成して返却する関数
def set_file_frame(parent_frame, label_text, file_type_list):
    file_frame = ttk.Frame(parent_frame)
    tk.Label(file_frame, text = label_text).pack(side = tk.LEFT)
    # テキストボックスの作成と配置
    file_frame.edit_box = tk.Entry(file_frame, width = 50)
    file_frame.edit_box.pack(side = tk.LEFT)
    # ボタンの作成と配置
    file_button = tk.Button(file_frame, text = 'FILE', width = 5\
        , command = lambda:open_file_command(file_frame.edit_box, file_type_list))
    file_button.pack(side = tk.LEFT)
    return file_frame
# フレームを作成する関数を呼び出して配置する関数
def set_main_frame(root_frame):
    # ファイル/フォルダ選択エリア作成
    file_set_frame_qraw = set_file_frame(root_frame, "アンケートデータ"\
        , [('アンケートデータ', '*.xlsx;*.csv'), ('他', '*')])
    file_set_frame_qraw.pack()
    file_set_frame_qlayout= set_file_frame(root_frame, "レイアウトデータ"\
        , [('レイアウトデータ', '*.xlsx;*.csv'), ('他', '*')])
    file_set_frame_qlayout.pack()
    dir_set_frame_qedit = set_dir_frame(root_frame, "出力先フォルダ")
    dir_set_frame_qedit.pack()
    # プルダウンエリア作成
    pulldown_frame = set_pulldown_frame(root_frame, "アンケートツール", ['Atcos', 'MApps', 'LINE RESEARCH PLATFORM'])
    pulldown_frame.pack()
    # 実行ボタンエリア作成
    run_frame = run_funct_frame(root_frame, file_set_frame_qraw, file_set_frame_qlayout, dir_set_frame_qedit, pulldown_frame, "加工実行")
    run_frame.pack()
#メイン
if __name__ == '__main__':
    root = tk.Tk()
    root.title("アンケートデータ 加工ツール")
    root.geometry("500x300")
    set_main_frame(root)
    root.mainloop()