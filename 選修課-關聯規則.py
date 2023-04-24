from google.colab import drive
drive.mount('/content/drive')

#匯入前記得要先到終端機 'pip install pandas'
import pandas as pd
import numpy as np
import csv

File_Path = '/content/drive/My Drive/關聯規則程式碼/修課ASR/'
df = pd.read_pickle(File_Path + '學生修課與成績資料_103to1102_20220317.pkl')

df.dropna(axis=0,inplace=True)

df["學制"]=df["虛學號"].str[0]
df["入學年"]=df["虛學號"].str[1:4].astype('int64')
df["院系組"]=df["系級代碼(成績)"].str[1:4]
df["級"]=df["系級代碼(成績)"].str[4]
df["系組名"]=df["系級代碼(成績)"].str[6:-2]
df["g"]=df["課程碼"].str[3]
df["g課名"]=df["g"]+df["課名"]

df=df[(df.學制=="A") & (df.入學年>=103) & (df.選別=="7選修") & 
      (~df["系級代碼(成績)"].str.contains("延")) & (~df.課名.str.contains("體育"))]
df.replace({'系組名': {"資管網科":"資管智網", "圖傳": "圖傳系"}},inplace=True)

df.drop_duplicates(subset=["虛學號","課名"],keep='first',inplace=True)

df["成績學年期"]=df['成績學年'].astype(str)+df['成績學期'].astype(str)

df.sort_values(["成績學年期"],
               axis = 0, ascending = True,
               inplace = True,
               na_position = "first")

df=df[["成績學年期","虛學號",'g課名','系組名']]

opopopop = pd.read_pickle(File_Path + 'G_ASR_資管資管.pkl')

import pickle
for d in d_lst: #26個系組
  tmp=[] #創建一個陣列
  df_d=df[df.系組名==d] #留下26個系組名的資料

  lst_stu=set(df_d["虛學號"]) #刪除陣列當中重複虛學號
  lst_stu=list(lst_stu) #把上面那行跑出來的東西變成list
  lst_stu.sort() #由小到大排列

  for s in lst_stu: #lst_stu是一個有很多 學號 的陣列
    df_stu=df_d[df_d.虛學號==s]
    if len(df_stu) > 0: #這個動作可以知道目前跑到的這個虛學號是不是在這個系裡面
      tmp.append(df_stu.g課名.tolist())

  with open(File_Path+"G_ASR_"+d+'.pkl', 'wb') as ff: #檔案不是已存在，這行程式碼可以創建新檔案
    pickle.dump(tmp, ff)

!pip install pyfpgrowth

import pyfpgrowth
# 資料  Q1

minS = 0.1 #支持度
minC = 0.7 #信賴度

for d in d_lst:
   with open(File_Path+"G_ASR_"+d+'.pkl', 'rb') as ff: #每個系的每個組 上面跑出來的結果(26個)
     Q1lst=pickle.load(ff)
     support = minS*len(Q1lst)
     # 獲取符合支援度規則資料
     patterns = pyfpgrowth.find_frequent_patterns(Q1lst, support) #支持度
     rules = pyfpgrowth.generate_association_rules(patterns, minC) #信賴度
     pd.DataFrame(list(patterns.items()),
                   columns=['g課名', '次數']).to_excel(File_Path+"GASR/G_Patterns_"+d+".xlsx")
     pd.DataFrame(list(rules.items()),
                   columns=['左', '右']).to_excel(File_Path+"GASR/G_Rules_"+d+".xlsx")