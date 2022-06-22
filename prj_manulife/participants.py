from cProfile import label
import os
from zipfile import BadZipfile
import numpy as np
import pandas as pd
from pathlib import Path

participants = {
    "ID" : list(),
    "Name" : list(),	
    "Gender" : list(),	
    "Address" : list(),	
    "Province" : list(),	
    "Email" : list(),	
    "Phone"	 : list(),
    "Phone_1" : list(),	
    "Phone_2" : list(),	
    "Phone_3" : list(),	
    "Phone_4" : list(),	
    "Phone_5" : list(),	
    "Link" : list(),	
    "Country" : list(),	
    "TransactionVIT" : list(),	
    "TransactionENG" : list(),	
    "Branch" : list(),	
    "Team" : list(),	
    "MassOrPremier" : list(),	
    "PRM" : list(),	
    "ProjectNameFilter"	: list(),
    "Filter1" : list(),	
    "Filter2" : list(),	
    "Filter3" : list()
}

df_participants = pd.DataFrame(data=participants, columns=list(participants.keys()))
df_participants.set_index(["ID"], inplace=True)

#1 sanji - 2 pauli - 3 first class - 4 nasty 5 newpher 6 niecece 7 kevin 8 newbie 9 uranus duoi 5 so 
nasty_columns = {
    "DUMMY_ID" : "ID",
    "AGT_NM" : "Name",
    "GENDER" : "Gender",
    "LOCATION" : "Province",
    "AGT_MOBILE" : "Phone", 
    "AGT_EMAIL" : "Email",
    "QUOTA" : "Filter1",
}

sanji_columns = {
    "DUMMY ID" : "ID",
    "CLI_NM" : "Name",
    "Khu vực" : "Province",
    "MOBL_PHON_NUM" : "Phone",
    "OTHR_PHON_NUM" : "Phone_1",
    "PRIM_PHON_NUM" : "Phone_2",
    "EMAIL_ADDR" : "Email"
}

uranus_columns = {
    "DUMMY ID" : "ID",
    "POLICY_OWNER_NAME" : "Name",
    "GENDER" : "Gender",
    "Handling Region" : "Province",
    "Phone3/Phone4" : "Address",
    "Phone 1" : "Phone",
    "Phone 2" : "Phone_1",
    "Phone 3" : "Phone_2",
    "Phone 4" : "Phone_3",
    "Quota" : "Filter1",
    "Handling Region" : "Filter2"
}

firstclass_columns = {
    "DUMMY ID" : "ID",
    "POLICY_OWNER_NAME" : "Name",
    "GENDER" : "Gender", 
    "LOCATION" : "Province",
    "CLIENT_REGION" : "Address",
    "MOBILE_NUMBER" : "Phone",
    "FIX_TELE_NUMBER" : "Phone_1",
    "OTHER_PHONE" : "Phone_2"
}

kevin_columns = {
    "DUMMY ID" : "ID",
    "AGT_NM" : "Name",
    "GENDER" : "Gender", 
    "LOCATION" : "Province",
    "CLIENT_REGION" : "Address",
    "MOBILE_NUMBER" : "Phone"
}

newbie_columns = {
    "DUMMY ID" : "ID",
    "POLICY_OWNER_NAME" : "Name",
    "GENDER" : "Gender", 
    "CLIENT_REGION" : "Province",
    "MOBILE_NUMBER" : "Phone",
    "FIX_TELE_NUMBER" : "Phone_1",
    "OTHER_PHONE" : "Phone_2",
    "OTHER_PHONE" : "Phone_3"
}

#Path
path = r"prj_manulife\DATABASE"
os.chdir(path)

writer = pd.ExcelWriter("../participant.xlsx", engine="xlsxwriter")



for dirname, dirs, files in os.walk(os.getcwd()):
    try:
        columns = np.empty
        project_name = ""

        entries = Path(dirname)

        prj_week = entries.name
        prj_month = entries.parent.name
        prj_year = entries.parent.parent.name

        for entry in entries.iterdir():
            if entry.is_file():
                df = pd.read_excel(entry, engine='openpyxl')
                
                if "NASTY" in entry.name:
                    project_name = "PRJ_NASTY"
                    link_final = "https://research3.ipsosinteractive.com/mriweb/mriweb.dll?i.project=S22000742"
                    columns = nasty_columns

                    uw_class_roles = {1 : "12E", 2 : "12", 3 : "18E", 4 : "18", 5 : "28", 6 : "21"}

                    df = df.loc[df["UW_CLASS"].apply(lambda x: str(x) in list(uw_class_roles.values()))]
                    
                    df["AGT_MOBILE"] = df["AGT_MOBILE"].apply(lambda x: "0{}".format(str(x)))
                    
                    df.loc[df["GENDER"] == "M", "GENDER"] = "Nam"
                    df.loc[df["GENDER"] == "F", "GENDER"] = "Nữ"

                    df.loc[df["LOCATION"] == "South", "LOCATION"] = "Hồ Chí Minh"
                    df.loc[df["LOCATION"] == "North", "LOCATION"] = "Hà Nội"

                    df.loc[df["QUOTA"] == 2, "QUOTA"] = "SaigonBank"
                """
                if "SANJI" in entry.name:
                    print(df)
                    project_name = "PRJ_SANJI"
                    link_final = "https://research3.ipsosinteractive.com/mriweb/mriweb.dll?i.project=S22000742"
                    columns = sanji_columns
                    
                    #dieu kien filter column CLI_NM va MOBL_PHON_NUM, OTHR_PHON_NUM, PRIM_PHON_NUM (4 tuan gan nhat)

                    df["DUMMY ID"] = df["DUMMY ID"].apply(lambda x: str(x))
                    df["MOBL_PHON_NUM"] = df["MOBL_PHON_NUM"].apply(lambda x: "{}".format(str(x).replace("'", "")) if len(x.replace("'", "")) > 0 else "")
                    df["OTHR_PHON_NUM"] = df["OTHR_PHON_NUM"].apply(lambda x: "{}".format(str(x).replace("'", "")) if len(x.replace("'", "")) > 0 else "")
                    df["PRIM_PHON_NUM"] = df["PRIM_PHON_NUM"].apply(lambda x: "{}".format(str(x).replace("'", "")) if len(x.replace("'", "")) > 0 else "")

                    df.loc[df["Khu vực"] == "HCM", "Khu vực"] = "Hồ Chí Minh"
                    df.loc[df["Khu vực"] == "HN", "Khu vực"] = "Hà Nội"
                if "FIRST CLASS" in filename:
                    project_name = "PRJ_FIRST_CLASS"
                    link_final = "https://research3.ipsosinteractive.com/mriweb/mriweb.dll?i.project=S22000742"
                    #columns = firstclass_columns
                if "KEVIN" in filename:
                    project_name = "PRJ_KEVIN"
                    link_final = "https://research3.ipsosinteractive.com/mriweb/mriweb.dll?i.project=S22000742"
                    #columns = kevin_columns
                """
                if columns is not np.empty:
                    df = df[list(columns.keys())]
                    df = df.rename(columns=columns)
                    df.set_index(["ID"], inplace=True)
                    
                    df["Address"] = df["Province"]
                    df["ProjectNameFilter"] = "{}_{}_{}{}".format(project_name, prj_week, prj_month, prj_year)
                    df["Link"] = link_final
                    df["Country"] = "Vietnam"

                    
                    df_participants = pd.concat([df_participants, df], ignore_index=False, axis=0)
    except BadZipfile as ex:
        print("File '% s' is not a zip file." % entry.name)
        
          
df_participants.to_excel(writer, sheet_name="Participant")
worksheet = writer.sheets["Participant"]

writer.save()



