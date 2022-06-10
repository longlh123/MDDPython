from email.policy import default
from inspect import stack
from ipaddress import v4_int_to_packed
import math
from turtle import left
from xml.dom.pulldom import ErrorHandler
import win32com.client as w32
from metadata import Metadata
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl

#MDM = w32.Dispatch('MDM.Document')
#MDM.Open(r'metadatas/PVN2021249_F61_W1_2022v1.mdd')

#questions = ["InstanceID","_Year","_Month","_Quarter","_ResProvinces","_City_Groups","_Target","_Users","_Class","_EPI_OOP_Inject","_Table_1","_Table_2","_Table_8","_Table_9","_Table_10_Before","_Table_10","_Table_13","_Table_14","_Table_15","_Table_16","_Table_17","_Table_18_Before","_Table_18","_Table_19","_Table_20","_Table_21","_Table_3","_UA1","_UA2","_UA3a","_UA3b","_UB1","_UB2","_UB4","_UB5","_Q3","_S16","_S16b","_S16c","_Q17","_UD3","_OP5c","_OP5e","_OP5d"]

questions = ["Respondent.ID","_Q1","_Phase"]

#try:
m = Metadata(r'metadatas/S22015945_DATA.mdd', r'metadatas/S22015945_DATA.ddf', questions)

df = m.convertToDataFrame() 
df = df.set_index("Respondent.ID")

obj_skus_define = {
    '_I1' : { 'variant' : 'Nội thất', 'segment' : 'Siêu cao cấp'},
    '_I1a' : { 'variant' : 'Nội thất', 'segment' : 'Siêu cao cấp'},
    '_I2' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I2a' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I3' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I3a' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I4a' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I4' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I5' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I6' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I7' : { 'variant' : 'Nội thất', 'segment' : 'Giá rẻ'},
    '_I8' : { 'variant' : 'Nội thất', 'segment' : 'Siêu cao cấp'},
    '_I8a' : { 'variant' : 'Nội thất', 'segment' : 'Siêu cao cấp'},
    '_I9' : { 'variant' : 'Nội thất', 'segment' : 'Cao cấp'},
    '_I10' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I11' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I12' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I13' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I14' : { 'variant' : 'Nội thất', 'segment' : 'Giá rẻ'},
    '_I15' : { 'variant' : 'Nội thất', 'segment' : 'Giá rẻ'},
    '_E1' : { 'variant' : 'Ngoại thất', 'segment' : 'Cao cấp'},
    '_E2' : { 'variant' : 'Ngoại thất', 'segment' : 'Siêu cao cấp'},
    '_E3' : { 'variant' : 'Ngoại thất', 'segment' : 'Cao cấp'},
    '_E4' : { 'variant' : 'Ngoại thất', 'segment' : 'Cao cấp'}, #_E4B
    '_E4M' : { 'variant' : 'Ngoại thất', 'segment' : 'Cao cấp'},
    '_E5' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_E6' : { 'variant' : 'Ngoại thất', 'segment' : 'Cao cấp'},
    '_E7' : { 'variant' : 'Ngoại thất', 'segment' : 'Siêu cao cấp'},
    '_E8' : { 'variant' : 'Ngoại thất', 'segment' : 'Cận cao cấp'},
    '_E9' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_E10' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_E11' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_W1' : { 'variant' : 'Chống thấm', 'segment' : 'Siêu cao cấp'},
    '_W3' : { 'variant' : 'Chống thấm', 'segment' : 'Cao cấp'},
    '_W2' : { 'variant' : 'Chống thấm', 'segment' : 'Cao cấp'},
    '_W4' : { 'variant' : 'Chống thấm', 'segment' : 'Cao cấp'},
    '_T1' : { 'variant' : 'Sơn dầu', 'segment' : 'Trung cấp'},
    '_T2' : { 'variant' : 'Sơn dầu', 'segment' : 'Trung cấp'},
    '_T3' : { 'variant' : 'Sơn dầu', 'segment' : 'Trung cấp'},
    '_I5e' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_I6e' : { 'variant' : 'Nội thất', 'segment' : 'Trung cấp'},
    '_E5e' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_E12e' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'},
    '_E13e' : { 'variant' : 'Ngoại thất', 'segment' : 'Trung cấp'}
}

cities = ["Hồ Chí Minh","Hà Nội", "Hải Phòng","Đà Nẵng", "Cần Thơ"]
segments = ["Siêu cao cấp", "Cao cấp", "Cận cao cấp", "Trung cấp", "Giá rẻ"]

df_data = df.loc[:, ["_Q1","_Phase[{_1}]._Q5[{_1}]._Q5_SKU","_Phase[{_1}]._Q5[{_2}]._Q5_SKU","_Phase[{_1}]._Q5[{_3}]._Q5_SKU","_Phase[{_1}]._Q5[{_4}]._Q5_SKU","_Phase[{_1}]._Q5[{_5}]._Q5_SKU","_Phase[{_1}]._Q5[{_6}]._Q5_SKU","_Phase[{_1}]._Q5[{_1}]._Q5_GiaTien","_Phase[{_1}]._Q5[{_2}]._Q5_GiaTien","_Phase[{_1}]._Q5[{_3}]._Q5_GiaTien","_Phase[{_1}]._Q5[{_4}]._Q5_GiaTien","_Phase[{_1}]._Q5[{_5}]._Q5_GiaTien","_Phase[{_1}]._Q5[{_6}]._Q5_GiaTien"]]

#BAR CHART - 100% STACKED COLUMNS CHART

#Combine multiple columns into a single dataframe 

df_data["Q5_SKU"] = df_data[df_data.columns[1:7]].apply(lambda x: ','.join(x.astype(str)), axis=1)
df_data["Q5_PRICE"] = df_data[df_data.columns[7:13]].apply(lambda x: ','.join(x.astype(str)), axis=1)

df_data.drop(df_data.columns[1:13], axis=1, inplace=True)

df_data.reset_index(-1, drop=False, inplace=True)
df_data.set_index(["Respondent.ID","_Q1"], drop=True, inplace=True, verify_integrity=False)
df2 = df_data.stack().str.split(',', expand=True).stack().unstack(-2).reset_index(-1, drop=True).reset_index()

print(df2.loc(["Q5_SKU"]).groupby(["Q5_SKU"]).count())


"""

obj_skus = {}

for i, r in df_data.iterrows():
    for i in range(1, 7):
        if math.isnan(r[i + 6]) is False:
            if r[i] not in obj_skus.keys():
                obj_skus[r[i]] = {}
            if r[0] not in obj_skus[r[i]].keys():
                obj_skus[r[i]][r[0]] = list()
            
            obj_skus[r[i]][r[0]].append(r[i + 6])

print(len(obj_skus.keys()))

fig, axs = plt.subplots(nrows=2, ncols=2, figsize = (10, 7)) 
fig.subplots_adjust(left=0.08, right=0.98, bottom=0.05, top=0.9, hspace=0.4, wspace=0.3)

arr_skus_selected = ["_I1","_I2","_I3","_I4"]

_row,_col = 0, 0

for sku in arr_skus_selected:
    d1 = []

    for k2, v2 in obj_skus[sku].items():
        if len(v2) > 0:
            d1.append(v2)
    
    axs[_row, _col].boxplot(d1, showmeans=True, meanline=True, sym="g o")
    axs[_row, _col].set_xticklabels(obj_skus[sku].keys())
    axs[_row, _col].set_title("{}".format(sku))
    _col += 1
    if _col == 2: 
        _col = 0
        _row += 1

plt.show()

"""
"""

fig, axs = plt.subplots(nrows=len(cities), ncols=len(segments), figsize = (15, 10))
fig.subplots_adjust(left=0.08, right=0.98, bottom=0.05, top=0.9, hspace=0.4, wspace=0.3)

_col, _row = 0, 0

for c in cities:
    _col = 0

    for s in segments:
        df_cities = df_data.query("_Q1 == '{}'".format(c))

        obj_skus = {}

        for index, row in df_cities.iterrows():
            for i in range(1, 7):
                if math.isnan(row[i + 6]) is False:
                    if obj_skus_define[row[i]]['segment'] == s:
                        if row[i] not in obj_skus.keys():
                            obj_skus[row[i]] = list()
                      ee  obj_skus[row[i]].append(row[i + 6]) 

        d1 = []

        for i in obj_skus:
            if len(obj_skus[i]) > 0:
                d1.append(obj_skus[i])
       
        axs[_row, _col].boxplot(d1, showmeans=True, meanline=True, sym="g o")
        axs[_row, _col].set_xticklabels(obj_skus.keys())
        axs[_row, _col].set_title("{} - {}".format(c, s))
        _col += 1 
    _row += 1

plt.savefig('plot.png')
writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name="Output")
worksheet = writer.sheets["Output"]
worksheet.insert_image('A1', 'plot.png')
writer.save()

plt.show()

"""











#df.to_excel(r'C:\Users\long.pham\Documents\MDDPython\export_dataframe.xlsx', index=False, header=True)
#except BaseException as ex:
#    print("Permission denied, file already open.")






