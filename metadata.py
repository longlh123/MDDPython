import win32com.client as w32
import pandas as pd

class Metadata:

    def __init__(self, mdd_path, ddf_path, questions):
        self.mdd_path = mdd_path
        self.ddf_path = ddf_path
        self.questions = questions

    def convertToDataFrame(self):
        MDM = w32.Dispatch('MDM.Document')
        MDM.Open(self.mdd_path)

        #column_names = list()

        #for question in self.questions:
        #    column_names.extend(self.__class__.getColumnNames(MDM, question))
        
        adoConn = w32.Dispatch('ADODB.Connection')

        conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(self.ddf_path, self.mdd_path)

        adoConn.Open(conn)

        adoRS = w32.Dispatch(r'ADODB.Recordset')
        adoRS.ActiveConnection = conn
        #adoRS.Open(r"SELECT " + ",".join(column_names) + " FROM VDATA")
        adoRS.Open(r"SELECT * FROM VDATA")

        d = {
            'columns' : list(),
            'values' : list()  
        }

        data = {}
        
        i = 0
        
        while not adoRS.EOF:

            r = self.getRow(MDM, adoRS)

            d['values'].append(r['values'])
            data[adoRS.Fields["Respondent.ID"].Value] = d['values'][0]

            if i == 0: d['columns'].append(r['columns']) 
            
            i += 1
            adoRS.MoveNext()

        MDM.Close()
        
        #Close and clean up
        adoRS.Close()
        adoRS = None

        adoConn.Close()
        adoConn = None

        return pd.DataFrame(data=d['values'], columns=d['columns'][0])
 
    def getRow(self, MDM, adoRS):
        r = {
            'columns' : list(),
            'values' : list()  
        }

        for question in self.questions:
            match self.__class__.objectTypeConstants(str(MDM.Fields[question].ObjectTypeValue)):
                case "mtVariable":
                    q = self.getValue(adoRS, MDM.Fields[question])
                    
                    r['values'].extend(q['values'])
                    r['columns'].extend(q['columns'])
                case "mtArray":
                    for variable in MDM.Fields[question].Variables:
                        q = self.getValue(adoRS, variable)

                        r['values'].extend(q['values'])
                        r['columns'].extend(q['columns'])
                case _:
                    r['values'].append(None)    
        return r

    def getValue(self, adoRS, question): 
        q = {
            'columns' : list(),
            'values' : list()  
        }
        
        max_range = 0
        
        column_name = question.FullName if self.__class__.objectTypeConstants(str(question.ObjectTypeValue)) != "mtVariable" else question.Variables[0].FullName
        
        if self.__class__.dataTypeConstants(question.DataType) == "mtCategorical":
            if question.Properties["py_showPunchingData"]:
                for category in question.Categories:
                    q['columns'].append("{}.{}".format(column_name, category.Name))
            else:
                max_range = question.MaxValue if question.MaxValue is not None else question.Categories.Count

                if question.MinValue == 1 and question.MaxValue == 1:
                    q['columns'].append(column_name)
                else:
                    for i in range(max_range):
                        q['columns'].append("{} ({}/{})".format(column_name, i + 1, max_range))
                        
            cats_resp = str(adoRS.Fields[column_name].Value)[1:(len(str(adoRS.Fields[column_name].Value))-1)].split(",")

            if question.Properties["py_showPunchingData"]:
                for category in question.Categories:
                    if category.Name in cats_resp:
                        q['values'].append(1)
                    else:

                        q['values'].append(0 if adoRS.Fields[column_name].Value != None else None)
            else:
                for i in range(max_range):
                    if i < len(cats_resp):
                        category = cats_resp[i]
                        
                        match question.Properties["py_showVariableValues"]:
                            case "Names":
                                q['values'].append(None if adoRS.Fields[column_name].Value == None else question.Categories[category].Name)
                            case "Labels":
                                q['values'].append(None if adoRS.Fields[column_name].Value == None else question.Categories[category].Label)
                            case _:
                                q['values'].append(None if adoRS.Fields[column_name].Value == None else int(category[1:len(category)]))
                    else:
                        q['values'].append(None)
        else:
            q['columns'].append(column_name)    
            q['values'].append(adoRS.Fields[column_name].Value)

        return q

    def objectTypeConstants(i):
        objectTypeConstants = {
            'ff' : 'mtUnknown',
            '0' : 'mtVariable', #Information, Text, Long, Double, Date, Categorical
            '1' : 'mtArray', #Loop
            '2' : 'mtGrid',
            '3' : 'mtClass', #Block Fields
            '4' : 'mtElement',
            '5' : 'mtElements',
            '6' : 'mtLabel',
            '7' : 'mtField',
            '8' : 'mtHelperFields',
            '9' : 'mtFields',
            'A' : 'mtTypes',
            'B' : 'mtProperties',
            'C' : 'mtRouting',
            'D' : 'mtContexts',
            'E' : 'mtLanguages',
            'F' : 'mtLevelObject',
            '10' : 'mtVariableInstance',
            '11' : 'mtRoutingItem',
            '12' : 'mtCompound',
            '13' : 'mtElementInstance',
            '14' : 'mtElementInstances',
            '15' : 'mtLanguage',
            '16' : 'mtRoutingItems',
            '17' : 'mtRanges',
            '18' : 'mtCategories',
            '19' : 'mtCategoryMap',
            '1A' : 'mtDataSources',
            '1B' : 'mtDocument',
            '1D' : 'mtVersion',
            '1E' : 'mtVersions',
            '1F' : 'mtVariables',
            '20' : 'mtDataSource',
            '21' : 'mtAliasMap',
            '22' : 'mtIndexElement',
            '23' : 'mtIndicesElements',
            '24' : 'mtPages',
            '25' : 'mtParameters',
            '26' : 'mtPage',
            '27' : 'mtItems',
            '28' : 'mtContext',
            '29' : 'mtContextAlternatives',
            '2A' : 'mtElementList',
            '2B' : 'mtGoto',
            '2C' : 'mtTemplate',
            '2D' : 'mtTemplates',
            '2E' : 'mtStyle',
            '2F' : 'mtNote',
            '30' : 'mtNotes',
            '31' : 'mtIfBlock',
            '32' : 'mtConditionalRouting',
            '33' : 'mtDBElements',
            '34' : 'mtDBQuestionDataProvider'
        }
        return objectTypeConstants.get(i, "Invalid object type value.")

    def dataTypeConstants(i):
        objDataTypeConstants = {
            0 : 'mtNone',
            1 : 'mtLong',
            2 : 'mtText',
            3 : 'mtCategorical',
            4 : 'mtObject',
            5 : 'mtDate',
            6 : 'mtDouble',
            7 : 'mtBoolean',
            8 : 'mtLevel'
        }
        return objDataTypeConstants.get(i, "Invalid data type constants.")