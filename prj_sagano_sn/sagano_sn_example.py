from cmath import nan
import sys
from unicodedata import name

from numpy import expand_dims
sys.path.append('c:\\Users\\long.pham\\Documents\\MDDPython')

from metadata import Metadata
import pandas as pd
import numpy as np
import re

project_name = 'prj_sagano_sn'
mdd_filename = 'metadatas/PVN2022038_SAGANO_SN_EXPORT_DAILY_v1.mdd'
ddf_filename = 'metadatas/PVN2022038_SAGANO_SN_EXPORT_DAILY_v1.ddf'

writer = pd.ExcelWriter("{}/output.xlsx".format(project_name), engine="xlsxwriter")

#Screener
questions = ["InstanceID","_S0a","_Gender","_S1","_AgeGroup","_S4","_Block_Kid_selected","_Q0_STC","_Q0_SBPS","_HouseHoldIncome","_Q15","_U1","_PE1","_PE2","_V1","_V2","_D1","_D2"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df.to_excel(writer, sheet_name="SCREENER")
worksheet = writer.sheets["SCREENER"]

#S7ab
questions = ["InstanceID","_S7a_Kid","_S7a_Mother","_S7b_Kids","_S7b_Mother"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for i in range(6):
    df["S7a_{}".format(i+1)] = df[df.columns[0:3]].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(df.columns[0:3], axis=1, inplace=True)
    
for i in range(6):
    df["S7b_{}".format(i+1)] = df[df.columns[0:3]].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(df.columns[0:3], axis=1, inplace=True)

df["S7a"] = df[df.columns[0:6]].apply(lambda x: '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)

df["S7b"] = df[df.columns[0:6]].apply(lambda x: '#'.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)

df_stack = df.stack().str.split("#", expand=True).rename(columns={0 : 'Kid_1',1 : 'Kid_2',2 : 'Kid_3',3 : 'Kid_4',4 : 'Kid_5',5 : 'Mother'}).stack().str.split(",", expand=True).rename(columns={0 : 'Sữa nước/ sữa tươi', 1 : 'Sữa trái cây', 2 : 'Sữa bột pha sẵn'}).stack().unstack(-3).reset_index().rename(columns={'level_1' : 'Kid_Mother_Level', 'level_2' : 'Brand_Level'})

df_stack.replace(to_replace=['None','nan'], value=np.nan, inplace=True)
df_stack.set_index(["InstanceID", "Kid_Mother_Level", "Brand_Level"], inplace=True)
df_stack.dropna(how='all', inplace=True)
df_stack.reset_index(inplace=True)
df_stack.set_index(["InstanceID"], inplace=True)

df_stack.to_excel(writer, sheet_name="S7AB")
worksheet = writer.sheets["S7AB"]

#Main
sn_brands = {
    0 : "1. Sữa VINAMILK Vinamilk (không nói rõ cụ thể nhãn hiệu nào)",
    1 : "2. Sữa tươi VINAMILK 100%/ VNM 100%",
    2 : "5. Sữa VINAMILK dạng bịch (gói/túi)",
    3 : "4. Sữa ADM/ ADM GOLD",
    4 : "7. Sữa VINAMILK Flex/ Flex",
    5 : "777. Sữa tươi Vinamilk 100% Organic",
    6 : "245. Sữa tươi Green Farm",
    7 : "12. Sữa Cô Gái Hà Lan/Dutch lady (không nói rõ cụ thể nhãn hiệu nào)",
    8 : "14. Sữa tươi Cô gái Hà Lan",
    9 : "13. Sữa Cô gái Hà Lan dạng bịch (gói/ túi)",
    10 : "83. Sữa Cô gái Hà Lan CAO KHỎE",
    11 : "84. Sữa Cô gái Hà Lan Organic/ Dutch Lady Organic",
    12 : "70. Sữa TH (không nói rõ cụ thể nhãn hiệu nào)",
    13 : "27. Sữa tươi TH/ TH True Milk",
    14 : "200. Sữa TH True Milk dạng bịch (gói/túi)",
    15 : "202. Sữa tươi TH Organic/ TH Topkid Organic",
    16 : "203. TH Yến Mạch",
    17 : "62. Sữa NUTIFOOD",
    18 : "42. Sữa Nutifood/Nuti/ Nutimilk dạng bịch (gói/túi)",
    19 : "65. Sữa tươi Nutimilk 100 điểm",
    20 : "139. Sữa IDP",
    21 : "213. Sữa LIF KUN/ Kun tươi vui",
    22 : "17. Sữa tươi Mộc Châu",
    23 : "31. Sữa Đà Lạt Milk",
    24 : "804. Sữa Nestle NutriStrong/ Sữa tươi Nestle/ Nestle Fresh Milk",
    25 : "971. Khác",
    26 : "972. Khác",
    27 : "973. Khác",
    28 : "974. Khác",
    29 : "975. Khác",
    30 : "976. Khác",
    31 : "977. Khác"
}

sbps_brands = {
    0 : "1. Vinamilk (không nói cụ thể sản phẩm nào)",
    1 : "2. Optimum Gold",
    2 : "3. Dielac Grow",
    3 : "4. Dielac Grow Plus/ Dielac Grow Plus Tổ Yến",
    4 : "5. Dielac Alpha Gold IQ",
    5 : "6. Yoko/ Yoko Gold",
    6 : "72. ColosGold",
    7 : "11. Nutifood (không nói cụ thể sản phẩm nào)",
    8 : "12. Nuti IQ Gold/ Nuti IQ Diamond",
    9 : "13. Nutifood Grow Plus/ Grow Plus diamond",
    10 : "14. Riso Opti Gold",
    11 : "15. Nuvita Grow/ Nuvita Grow Diamond",
    12 : "18. Pedia Plus",
    13 : "19. Nuvi Grow",
    14 : "21. Cô Gái Hà Lan/ Friso Gold (không nói cụ thể sản phẩm nào)",
    15 : "31. Abbott (không nói cụ thể sản phẩm nào)",
    16 : "32. Abbott Grow Gold",
    17 : "33. Pedia Sure",
    18 : "34. Similac",
    19 : "41. Mead Johnson/ ENFA A+ (không nói cụ thể sản phẩm nào)",
    20 : "51. Nestle/ NAN (không nói cụ thể sản phẩm nào)",
    21 : "60. Vita Dairy",
    22 : "61. Colosbaby",
    23 : "62. Oggi Gold",
    24 : "70. Nutricare",
    25 : "71. Metacare",
    26 : "971. Khác",
    27 : "972. Khác",
    28 : "973. Khác",
    29 : "974. Khác",
    30 : "975. Khác",
    31 : "976. Khác",
    32 : "977. Khác"
}

sdn_brands = {
    0 : "1. Fami",
    1 : "2. Vinasoy",
    2 : "3. Nuti",
    3 : "6. Vinamilk",
    4 : "10. Goldsoy",
    5 : "14. Homesoy",
    6 : "15. Soy Secretz",
    7 : "16. Number 1 Soya",
    8 : "971. Khác",
    9 : "972. Khác",
    10 : "973. Khác",
    11 : "974. Khác",
    12 : "975. Khác",
    13 : "976. Khác",
    14 : "977. Khác"
}
 
#SN KID
questions = ["InstanceID","_Q1a","_Q1ab_MERGER","_Q2_TOTAL","_Q3","_Q4","_Q5","_Q6","_CONSIDER","_FLAGCS","_SWTEN","_PBVC","_CLBVC"]

question_names = ("TOM","SPONTANOUS","AIDED_AWARENESS","EVER_USED","P1M","BUMO","PREVIOUS_BUMO","CONSIDERATION","CONSIDERATION_SET","SWTEN","PBVC","CLBVC")

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_main = df.stack().str.split(',', expand=True).rename(columns=sn_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_main.set_index(["InstanceID","Brand_Level"], inplace=True)

#SN KID QI
questions = ["InstanceID","_QI"]

qi_attrs = ('_1','_4','_5','_8','_10','_12','_13','_16','_17','_18','_19','_20','_21','_24','_25','_26','_27')

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in qi_attrs:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == "QI.{}".format(q)]

    df["QI_ATTR{}".format(q)] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_qi = df.stack().str.split(',', expand=True).rename(columns=sn_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_qi.set_index(["InstanceID","Brand_Level"], inplace=True)

#SN KID Q14

questions = ["InstanceID","_Q14"]

q14_attrs = ('_1','_3','_4','_7','_5','_6','_10','_11')

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in q14_attrs:
    cols = list(filter(lambda x: re.match('Q14\.(.*)\.({})$'.format(q), x), df.columns))

    df["Q14_ATTR{}".format(q)] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_q14 = df.stack().str.split(',', expand=True).rename(columns=sn_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_q14.set_index(["InstanceID","Brand_Level"], inplace=True)

df_merge_kid = pd.concat([df_main, df_qi], ignore_index=False, sort=False, axis=1)
df_merge_kid = pd.concat([df_merge_kid, df_q14], ignore_index=False, sort=False, axis=1)
df_merge_kid.reset_index(inplace=True)
df_merge_kid["Variant_Level"] = df_merge_kid.apply(lambda x: "SN_KID", axis=1)
df_merge_kid.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

#SN MOTHER

questions = ["InstanceID","_Q1a","_Q1ab_MERGER","_Q2_TOTAL","_M3","_M4","_M5","_M6","_M7"]

question_names = ("TOM","SPONTANOUS","AIDED_AWARENESS","EVER_USED","P1M","BUMO","PREVIOUS_BUMO","CONSIDERATION")

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_merge_mother = df.stack().str.split(',', expand=True).rename(columns=sn_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'Brand_Level'})
df_merge_mother["Variant_Level"] = df_merge_mother.apply(lambda x: "SN_MOTHER", axis=1)
df_merge_mother.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

df_merge = pd.concat([df_merge_kid, df_merge_mother], ignore_index=False, axis=0)

#SBPS
questions = ["InstanceID","_Q1a_SBPS","_Q1ab_SBPS_MERGER","_Q2_SBPS_TOTAL","_Q3_SBPS","_Q4_SBPS","_Q5_SBPS","_Q6_SBPS","_Q7_SBPS"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df.dropna(how='all', inplace=True)

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_sbps = df.stack().str.split(',', expand=True).rename(columns=sbps_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : "Brand_Level"}).set_index(["InstanceID"])

df_sbps["Variant_Level"] = df_sbps.apply(lambda x: "SBPS", axis=1)
df_sbps.reset_index(inplace=True)
df_sbps.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

df_merge = pd.concat([df_merge, df_sbps], ignore_index=False, axis=0)

#SDN
questions = ["InstanceID","_Q1a_SDN","_Q1ab_SDN_MERGER","_Q2_SDN_TOTAL","_Q3_SDN","_Q4_SDN","_Q5_SDN","_Q6_SDN","_Q7_SDN"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df.dropna(how='all', inplace=True)

for q in question_names:
    cols = [col for col in df.columns if col[0:(len(col) if len(col.split('.')) == 0 else col.rfind('.'))] == q]

    df[q] = df[cols].apply(lambda x: ','.join(x.astype(str)), axis=1)
    df.drop(cols, axis=1, inplace=True)

df_sdn = df.stack().str.split(',', expand=True).rename(columns=sdn_brands).stack().unstack(-2).reset_index().rename(columns={'level_1' : "Brand_Level"})

df_sdn["Variant_Level"] = df_sdn.apply(lambda x: "SDN", axis=1)
df_sdn.reset_index(inplace=True)
df_sdn.set_index(["InstanceID","Variant_Level","Brand_Level"], inplace=True)

df_merge = pd.concat([df_merge, df_sdn], ignore_index=False, axis=0)
df_merge.reset_index(inplace=True)
df_merge.set_index(["InstanceID"], inplace=True)

df_merge.to_excel(writer, sheet_name="BRANDS")
worksheet = writer.sheets["BRANDS"]

#TVC
questions = ["InstanceID","_T1a"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df["TVC_Level"] = df[df.columns[0:6]].apply(lambda x : ','.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)
df["T1a"] = df[df.columns[0:6]].apply(lambda x : ','.join(x.astype(str)), axis=1)
df.drop(df.columns[0:6], axis=1, inplace=True)

df_2 = df.stack().str.split(',', expand=True).stack().unstack(-2).reset_index(-1, drop=True).reset_index().set_index(["InstanceID","TVC_SELECTED"])

questions = ["InstanceID","_ST_TVC"]

m = Metadata(r'{}/{}'.format(project_name, mdd_filename), r'{}/{}'.format(project_name, ddf_filename), questions)

df = m.convertToDataFrame() 
df = df.set_index("InstanceID")

df["TVC_Level"] = df[df.columns[0:2]].apply(lambda x : ','.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)
df["ST3"] = df[df.columns[0:2]].apply(lambda x : ','.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)
df.drop(df.columns[0:2], axis=1, inplace=True)
df["ST5"] = df[df.columns[0:2]].apply(lambda x : ','.join(x.astype(str)), axis=1)
df.drop(df.columns[0:2], axis=1, inplace=True)

df_3 = df.stack().str.split(',', expand=True).stack().unstack(-2).reset_index().rename(columns={'level_1' : 'TVC_ROT'})
df_4 = df_3[df_3['TVC_Level'] != 'None']
df_4.set_index(["InstanceID","TVC_Level"], inplace=True)

df_merge = pd.merge(df_2, df_4, left_index=True, right_index=True, how="outer")
df_merge.reset_index(-1, inplace=True)

df_merge.to_excel(writer, sheet_name="TVC")
worksheet = writer.sheets["TVC"]

writer.save()
