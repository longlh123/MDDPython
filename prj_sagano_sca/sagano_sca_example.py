from cmath import nan
import sys
from unicodedata import name

from numpy import expand_dims
sys.path.append('c:\\Users\\long.pham\\Documents\\MDDPython')

from metadata import Metadata
import pandas as pd
import numpy as np
import re

project_name = 'prj_sagano_sca'
mdd_filename = 'metadatas/SAGANO_SCA_DATA_DAILY_v1.mdd'
ddf_filename = 'metadatas/SAGANO_SCA_DATA_DAILY_v1.ddf'

writer = pd.ExcelWriter("{}/output.xlsx".format(project_name), engine="xlsxwriter")

#Screener
questions = ["InstanceID","_S0a","_S1","_AgeGroup","_S4","_Kid_AgeGroup","_Kid_Gender","_HouseHoldIncome","_Q15","_U1","_PE1","_PE2","_D1","_D2"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df.to_excel(writer, sheet_name="SCREENER")
worksheet = writer.sheets["SCREENER"]

#S7ab
questions = ["InstanceID","_S7a_Kid","_S7a","_S7b_Kids","_S7b"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for i in range(6):
    df["S7a_{}".format(i+1)] = df[df.columns[0:2]].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(df.columns[0:2], axis=1, inplace=True)
    
for i in range(6):
    df["S7b_{}".format(i+1)] = df[df.columns[0:2]].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(df.columns[0:2], axis=1, inplace=True)

df["S7a"] = df[df.columns[0:6]].apply(lambda x: '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)

df["S7b"] = df[df.columns[0:6]].apply(lambda x: '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)

df_stack = df.stack().str.split("#", expand=True).rename(columns={0 : 'Kid_1',1 : 'Kid_2',2 : 'Kid_3',3 : 'Kid_4',4 : 'Kid_5',5 : 'Mother'}).stack().str.split(",", expand=True).rename(columns={0 : 'Sữa chua ăn', 1 : 'Sữa trái cây'}).stack().unstack(-3).reset_index().rename(columns={'level_1' : 'Kid_Mother_Level', 'level_2' : 'Brand_Level'})

df_stack.replace(to_replace=['None','nan'], value=np.nan, inplace=True)
df_stack.set_index(["InstanceID", "Kid_Mother_Level", "Brand_Level"], inplace=True)
df_stack.dropna(how='all', inplace=True)
df_stack.reset_index(inplace=True)
df_stack.set_index(["InstanceID"], inplace=True)

df_stack.to_excel(writer, sheet_name="S7AB")
worksheet = writer.sheets["S7AB"]

#Main
sca_brands = {
    0 :  "1. Sữa chua ăn Vinamilk (nói chung)",
    1 :  "200. Vinamilk trắng /có đường",
    2 :  "6. Vinamilk nha đam",
    3 :  "129. Vinamilk Trái cây (Dâu/ Trái cây/ Dứa/ Lựu)",
    4 :  "12. Vinamilk Susu/ Susu",
    5 :  "80. Vinamilk Star/ Star",
    6 :  "209. Vinamilk Love Yogurt (Trắng/ nếp cẩm/ trân châu đường đen)/ Love Yogurt",
    7 :  "20. Sữa chua ăn Cô Gái Hà Lan/ Dutch Lady",
    8 :  "34. Sữa chua ăn TH/ TH/ TH true yogurt",
    9 :  "35. Sữa chua ăn TH Topkid",
    10 : "212. Sữa chua ăn Ba Vì/ Lif/ Love’in farm",
    11 : "213. Sữa chua ăn Kun",
    12 : "93. Sữa chua ăn Nuti/ Nutimilk",
    13 : "54. Sữa chua ăn Mộc Châu/ Mộc Châu nếp cẩm",
    14 : "215. Sữa chua ăn Hoff",
    15 : "94. Sữa chua ăn Kidsmix",
    16 : "7. Vinamilk 100%",
    17 : "8. Vinamilk Probi/ Probi",
    18 : "206. Sữa chua ăn Saigon Milk/Sg Milk",
    19 : "23. Sữa chua ăn Welyo",
    20 : "971. Khác",
    21 : "972. Khác",
    22 : "973. Khác",
    23 : "974. Khác",
    24 : "975. Khác",
    25 : "976. Khác",
    26 : "977. Khác"
}

stc_brands = {
    0 : "1. Vinamilk (nói chung)",
    1 : "2. Hero",
    2 : "4. Kun",
    3 : "8. Nutriboost",
    4 : "9. Tropicana",
    5 : "14. Nuvi",
    6 : "971. Khác",
    7 : "972. Khác",
    8 : "973. Khác",
    9 : "974. Khác",
    10 : "975. Khác",
    11 : "976. Khác",
    12 : "977. Khác"
}

#SCA KID
questions = ["InstanceID","_Q1a","_Q1ab","_Q2_TOTAL","_Q3","_Q4","_Q5","_Q6","_CONSIDER","_FLAGCS","_SWTEN","_PBVC","_CLBVC"]

question_names = ("TOM","SPONTANOUS","AIDED_AWARENESS","EVER_USED","P1M","BUMO","PREVIOUS_BUMO","CONSIDERATION","CONSIDERATION_SET","SWTEN","PBVC","CLBVC")

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_main = df.stack().str.split(',', expand=True).rename(columns=sca_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_main.set_index(["InstanceID","Brand_Level"], inplace=True)

#SCA KID QI
questions = ["InstanceID","_QI"]

qi_attrs = ('_5','_6','_7','_13','_14','_18','_22','_26','_28','_29','_33','_34','_35','_36','_37')

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in qi_attrs:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == "QI.{}".format(q)]

    df["QI_ATTR{}".format(q)] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_qi = df.stack().str.split(',', expand=True).rename(columns=sca_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_qi.set_index(["InstanceID","Brand_Level"], inplace=True)

#SCA KID Q14

questions = ["InstanceID","_Q14"]

q14_attrs = ('_1','_3','_4','_7','_5','_6','_10','_11')

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in q14_attrs:
    cols = list(filter(lambda x: re.match('Q14\.(.*)\.({})$'.format(q), x), df.columns))

    df["Q14_ATTR{}".format(q)] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_q14 = df.stack().str.split(',', expand=True).rename(columns=sca_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_q14.set_index(["InstanceID","Brand_Level"], inplace=True)

df_merge_kid = pd.concat([df_main, df_qi], ignore_index=False, sort=False, axis=1)
df_merge_kid = pd.concat([df_merge_kid, df_q14], ignore_index=False, sort=False, axis=1)
df_merge_kid.reset_index(inplace=True)
df_merge_kid["Variant_Level"] = df_merge_kid.apply(lambda x: "SCA_KID", axis=1)
df_merge_kid.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

#SCA MOTHER

questions = ["InstanceID","_Q1a","_Q1ab","_Q2_TOTAL","_M3","_M4","_M5","_M6","_M7"]

question_names = ("TOM","SPONTANOUS","AIDED_AWARENESS","EVER_USED","P1M","BUMO","PREVIOUS_BUMO","CONSIDERATION")

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_merge_mother = df.stack().str.split(',', expand=True).rename(columns=sca_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_merge_mother["Variant_Level"] = df_merge_mother.apply(lambda x: "SCA_MOTHER", axis=1)
df_merge_mother.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

df_merge = pd.concat([df_merge_kid, df_merge_mother], ignore_index=False, axis=0)

#STC
questions = ["InstanceID","_Q1a_STC","_Q1ab_STC","_Q2_STC_Total","_Q3_STC","_Q4_STC","_Q5_STC","_Q6_STC","_BC_STC"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_stc = df.stack().str.split(',', expand=True).rename(columns=stc_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : "Brand_Level"}).set_index(["InstanceID"])

df_stc["Variant_Level"] = df_stc.apply(lambda x: "STC", axis=1)
df_stc.reset_index(inplace=True)
df_stc.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

df_merge = pd.concat([df_merge, df_stc], ignore_index=False, axis=0)
df_merge.reset_index(inplace=True)
df_merge.set_index(["InstanceID"], inplace=True)

df_merge.to_excel(writer, sheet_name="BRANDS")
worksheet = writer.sheets["BRANDS"]

#TVC
questions = ["InstanceID","_T1a"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df["TVC_SELECTED"] = df[df.columns[0:2]].apply(lambda x : '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)
df["T1a"] = df[df.columns[0:2]].apply(lambda x : '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)

df_2 = df.stack().str.split('#', expand=True).stack().unstack(-2).reset_index(-1, drop=True).reset_index().set_index(["InstanceID","TVC_SELECTED"])

questions = ["InstanceID","_ST_TVC"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df["TVC_SELECTED"] = df[df.columns[0:2]].apply(lambda x : '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)
df["ST3"] = df[df.columns[0:2]].apply(lambda x : '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)
df.drop(df.columns[0:2], axis=1, inplace=True)
df["ST5"] = df[df.columns[0:2]].apply(lambda x : '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)

df_3 = df.stack().str.split('#', expand=True).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'TVC_ROT'})
df_4 = df_3[df_3['TVC_SELECTED'] != 'None']
df_4.set_index(["InstanceID","TVC_SELECTED"], inplace=True)



df_merge = pd.merge(df_2, df_4, left_index=True, right_index=True, how="outer")
df_merge.reset_index(-1, inplace=True)

df_merge.to_excel(writer, sheet_name="TVC")
worksheet = writer.sheets["TVC"]

writer.save()
