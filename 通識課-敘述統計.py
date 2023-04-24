import pandas as pd
File_Path = '/content/drive/My Drive/Master/Analyze Data/'
Result_Path = '/content/drive/My Drive/Master/Colab Code/0605(108-109通識重作)/'

from google.colab import drive
drive.mount('/content/drive')

Class_Teacher = pd.read_excel(File_Path + '學期開課時段及授課老師.xlsx', sheet_name = '學期開課及授課老師')
Class_108_12 = pd.read_excel(File_Path + '學生預選課表_108全.xlsx', sheet_name = '學生課表108')
Class_109_1 = pd.read_excel(File_Path + '學生預選課表_109上.xlsx', sheet_name = '學生課表109上')

Class_僅學士班 = Class_Teacher.loc[(Class_Teacher["開課學制"] == 'A學士班') & 
         (Class_Teacher['選上人數'] >= 20) & 
         (Class_Teacher['星期'] != 0)  (Class_Teacher["星期"].isnull() != True)]

Class_僅學士班

All_Classes = Class_108_12.append(Class_109_1)
All_Classes

All_Classes_僅通識課 = All_Classes.loc[(All_Classes['選別碼'] == '9通識') &
            (All_Classes['興趣志願碼'].isnull() != True)]
All_Classes_僅通識課

通識_Merge_學士班 = Class_僅學士班.merge(All_Classes_僅通識課, on = ['學年','學期','課碼']) #兩個檔案合併時，學年,學期,課碼會合併再一起
通識_Merge_學士班

Teacher_刪除重複 = 通識_Merge_學士班.drop_duplicates(subset=['學年','學期','課碼','虛學號','課名_x'], keep='first')

#-------------------下半部為預選人數

只要108_1 = Teacher_刪除重複.loc[(Teacher_刪除重複['學年'] == 108) & (Teacher_刪除重複['學期'] == 1)]

Group_1081選的人數 = 只要108_1[['學年','學期','課碼','虛學號','課名_x','興趣領域名']]

Result_計算人數1081 = Group_1081選的人數.groupby(['學年','學期','課碼','課名_x','興趣領域名']).count()

Result_計算人數1081.sort_values('虛學號', ascending = False, inplace = True) #ascending為升冪(由小到大排序)，inplace表示直接修改原始資料

Result_計算人數1081.to_excel(Result_Path + '1081預選人數.xlsx', merge_cells = False)

只要108_2 = Teacher_刪除重複.loc[(Teacher_刪除重複['學年'] == 108) & (Teacher_刪除重複['學期'] == 2)]

Group_1082選的人數 = 只要108_2[['學年','學期','課碼','虛學號','課名_x','興趣領域名']]

Result_計算人數1082 = Group_1082選的人數.groupby(['學年','學期','課碼','課名_x','興趣領域名']).count()

Result_計算人數1082.sort_values('虛學號', ascending = False, inplace = True)

Result_計算人數1082.to_excel(Result_Path + '1082預選人數.xlsx', merge_cells = False)

只要109_1 = Teacher_刪除重複.loc[(Teacher_刪除重複['學年'] == 109) & (Teacher_刪除重複['學期'] == 1)]

Group_1091選的人數 = 只要109_1[['學年','學期','課碼','虛學號','課名_x','興趣領域名']]

Result_計算人數1091 = Group_1091選的人數.groupby(['學年','學期','課碼','課名_x','興趣領域名']).count() #因為要計算人數，所以透過虛學號欄位進行

Result_計算人數1091.sort_values('虛學號', ascending = False, inplace = True) #ascending為升冪(由小到大排序)，inplace表示直接修改原始資料

Result_計算人數1091.to_excel(Result_Path + '1091預選人數.xlsx', merge_cells = False)

#-------------------下半部為志願碼

Group_1081興趣志願碼 = 只要108_1[['學年','學期','課碼','興趣志願碼','課名_x','興趣領域名']]

Result_興趣平均分數1081 = Group_1081興趣志願碼.groupby(['學年','學期','課碼','課名_x','興趣領域名']).mean()

Result_興趣平均分數1081.sort_values('興趣志願碼', ascending = True, inplace = True)

Result_興趣平均分數1081.to_excel(Result_Path + '1081志願碼.xlsx', merge_cells = False)

Group_1082興趣志願碼 = 只要108_2[['學年','學期','課碼','興趣志願碼','課名_x','興趣領域名']]

Result_興趣平均分數1082 = Group_1082興趣志願碼.groupby(['學年','學期','課碼','課名_x','興趣領域名']).mean()

Result_興趣平均分數1082.sort_values('興趣志願碼', ascending = True, inplace = True)

Result_興趣平均分數1082.to_excel(Result_Path + '1082志願碼.xlsx', merge_cells = False)

Group_1091興趣志願碼 = 只要109_1[['學年','學期','課碼','興趣志願碼','課名_x','興趣領域名']]

Result_興趣平均分數1091 = Group_1091興趣志願碼.groupby(['學年','學期','課碼','課名_x','興趣領域名']).mean()

Result_興趣平均分數1091.sort_values('興趣志願碼', ascending = True, inplace = True)

Result_興趣平均分數1091.to_excel(Result_Path + '1091志願碼.xlsx', merge_cells = False)

#------------------------下半部為合併後輸出

Final_1081輸出 = Result_計算人數1081.merge(Result_興趣平均分數1081, on = ['學年','學期','課碼','課名_x','興趣領域名'])
Final_1081輸出.to_excel(Result_Path + '1081預選+志願碼.xlsx', merge_cells = False)

Final_1082輸出 = Result_計算人數1082.merge(Result_興趣平均分數1082, on = ['學年','學期','課碼','課名_x','興趣領域名'])
Final_1082輸出.to_excel(Result_Path + '1082預選+志願碼.xlsx', merge_cells = False)

Final_1091輸出 = Result_計算人數1091.merge(Result_興趣平均分數1091, on = ['學年','學期','課碼','課名_x','興趣領域名'])
Final_1091輸出.to_excel(Result_Path + '1091預選+志願碼.xlsx', merge_cells = False)