"""
从文档中提取auto刀阵容并生成excel表格文件
环境依赖：
python >=3.6
pandas
openpyxl
StyleFrame
Jinja2
"""
import pandas as pd
import numpy as np

Filename = '一阶段作业收集.xlsx'

df = pd.read_excel(Filename)

# 提取所有含auto刀作业的行
auto = df[df['Unnamed: 1'].str.contains('at|bt|ct', na=False,case=False)]
auto2 = auto.copy()
auto2['Unnamed: 3'] = auto2['Unnamed: 3'].fillna('null') #Give NaN a value Null
auto2 = auto2[~auto2['Unnamed: 3'].isin(['null'])] # Delete the row containing Null

# 判断半auto刀作业并标记
auto_1 = auto2[auto2['Unnamed: 1'].str.contains('t1', na=False,case=False)].iloc[:, 1:13]
for row in range(auto2.shape[0]):
    flag1 = 0
    for column in range(auto2.shape[1]):
        if flag1 == 1:
            continue
        else:
#             print(type(auto2.iloc[row,column]))
            if type(auto2.iloc[row,column]) == str:
                if '半auto' in auto2.iloc[row,column]:
                    newstr = '['+auto2.iloc[row,1]+']'
#                     print(newstr)
                    auto2.iloc[row,1] = newstr
                    flag1 = 1
                elif '半AUTO' in auto2.iloc[row,column]:
                    newstr = '['+auto2.iloc[row,1]+']'
#                     print(newstr)
                    auto2.iloc[row,1] = newstr
                    flag1 = 1
                else:
                    continue

# 生成各个王作业的DataFrame
auto_1 = auto2[auto2['Unnamed: 1'].str.contains('t1', na=False,case=False)].iloc[:, 1:8]
auto_2 = auto2[auto2['Unnamed: 1'].str.contains('t2', na=False,case=False)].iloc[:, 1:8]
auto_3 = auto2[auto2['Unnamed: 1'].str.contains('t3', na=False,case=False)].iloc[:, 1:8]
auto_4 = auto2[auto2['Unnamed: 1'].str.contains('t4', na=False,case=False)].iloc[:, 1:8]
auto_5 = auto2[auto2['Unnamed: 1'].str.contains('t5', na=False,case=False)].iloc[:, 1:8]

# 判断x面作业
phase = None
if 'c' in auto_1.iloc[0,0]:
    phase = 'C'
elif 'b' in auto_1.iloc[0,0]:
    phase = 'B'
elif 'a' in auto_1.iloc[0,0]:
    phase = 'A'

# 构造生成excel表格的总体DataFrame
df_auto1 = pd.DataFrame(auto_1.values, index = range(auto_1.shape[0]))
df_auto2 = pd.DataFrame(auto_2.values, index = range(auto_2.shape[0]))
df_auto3 = pd.DataFrame(auto_3.values, index = range(auto_3.shape[0]))
df_auto4 = pd.DataFrame(auto_4.values, index = range(auto_4.shape[0]))
df_auto5 = pd.DataFrame(auto_5.values, index = range(auto_5.shape[0]))
df_auton = pd.DataFrame(np.array([['']*auto_5.shape[1]]*auto_5.shape[0],dtype=object))

index_1 = pd.DataFrame([[phase+'1']+['']*6])
index_2 = pd.DataFrame([[phase+'2']+['']*6])
index_3 = pd.DataFrame([[phase+'3']+['']*6])
index_4 = pd.DataFrame([[phase+'4']+['']*6])
index_5 = pd.DataFrame([[phase+'5']+['']*6])

df_auto135 = pd.concat([index_1,df_auto1,index_3,df_auto3,index_5,df_auto5],
                       axis=0,ignore_index=True)
df_auto24 = pd.concat([index_2,df_auto2,index_4,df_auto4],
                       axis=0,ignore_index=True)
df_autoall = pd.concat([df_auto135,df_auto24],
                       axis=1,ignore_index=True)

# 为生成的DataFrame附加样式
from styleframe import StyleFrame,Styler,utils

auto_sf = StyleFrame(df_autoall)
index_auto = [auto_1.shape[0]+1,auto_2.shape[0]+1,
              auto_1.shape[0]+auto_3.shape[0]+2,auto_2.shape[0]+auto_4.shape[0]+2,
              auto_1.shape[0]+auto_3.shape[0]+auto_5.shape[0]+3]
column_auto = [list(range(auto_1.shape[1])),
               list(range(auto_1.shape[1],auto_1.shape[1]+auto_2.shape[1])),
               list(range(auto_1.shape[1])),
               list(range(auto_1.shape[1],auto_1.shape[1]+auto_2.shape[1])),
               list(range(auto_1.shape[1]))]

# 设置单元格颜色
# 一王
auto_sf.apply_style_by_indexes(auto_sf.index[0],
                               styler_obj = Styler(bg_color= 'FFFFC000'),
                               cols_to_style = column_auto[0])
auto_sf.apply_style_by_indexes(auto_sf.index[1:index_auto[0]],
                               styler_obj = Styler(bg_color= 'FFFFE1B2'),
                               cols_to_style = 0)
# 二王
auto_sf.apply_style_by_indexes(auto_sf.index[0],
                               styler_obj = Styler(bg_color= 'FFA9D08E'),
                               cols_to_style = column_auto[1])
auto_sf.apply_style_by_indexes(auto_sf.index[1:index_auto[1]],
                               styler_obj = Styler(bg_color= 'FFD9EAD3'),
                               cols_to_style = 7)
#三王
auto_sf.apply_style_by_indexes(auto_sf.index[index_auto[0]],
                               styler_obj = Styler(bg_color= 'FF9BC2E6'),
                               cols_to_style = column_auto[2])
auto_sf.apply_style_by_indexes(auto_sf.index[(index_auto[0]+1):index_auto[2]],
                               styler_obj = Styler(bg_color= 'FFDFF8FF'),
                               cols_to_style = 0)
#四王
auto_sf.apply_style_by_indexes(auto_sf.index[index_auto[1]],
                               styler_obj = Styler(bg_color= 'FFCC99FF'),
                               cols_to_style = column_auto[3])
auto_sf.apply_style_by_indexes(auto_sf.index[(index_auto[1]+1):index_auto[3]],
                               styler_obj = Styler(bg_color= 'FFCFC7F4'),
                               cols_to_style = 7)
#五王
auto_sf.apply_style_by_indexes(auto_sf.index[index_auto[2]],
                               styler_obj = Styler(bg_color= 'FFFF99FF'),
                               cols_to_style = column_auto[4])
auto_sf.apply_style_by_indexes(auto_sf.index[(index_auto[2]+1):index_auto[4]],
                               styler_obj = Styler(bg_color= 'FFFEE4FF'),
                               cols_to_style = 0)

#设置字体
for index, row in df_autoall.iterrows():
    for  column, value in row.iteritems(): 
        color = auto_sf.loc[index,column].style.bg_color
        auto_sf.apply_style_by_indexes(auto_sf[auto_sf[column].isin(['克'])],
                              styler_obj=Styler(#bg_color = color, 
                                                bold = True,font_color='FF70AD47'),
                              cols_to_style= column)
        color = auto_sf.loc[index,column].style.bg_color
        auto_sf.apply_style_by_indexes(auto_sf[auto_sf[column].isin(['狼'])],
                              styler_obj=Styler(#bg_color = color,
                                                bold = True,font_color='FF4472C4'),
                              cols_to_style= column)
        color = auto_sf.loc[index,column].style.bg_color
        auto_sf.apply_style_by_indexes(auto_sf[auto_sf[column].isin(['吉塔'])],
                              styler_obj=Styler(#bg_color = 'FFEEEEEE',
                                                bold = True,font_color='FFED7D31'),
                              cols_to_style= column)
        color = auto_sf.loc[index,column].style.bg_color
        auto_sf.apply_style_by_indexes(auto_sf[auto_sf[column].isin(['病娇'])],
                              styler_obj=Styler(#bg_color = 'FFEEEEEE',
                                                bold = True,font_color='FF7030A0'),
                              cols_to_style= column)
    

#保存文件
ew = StyleFrame.ExcelWriter('auto_test.xlsx')
auto_sf.to_excel(ew)
ew.save()
ew.close()