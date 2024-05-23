from ProcoreBaseFunction import *

ProcoreDataDict={}
def procoreTemplateDataRead(procoreExportFile, LATEST_COST_CODE_File):
    #####procore template data read
    map_template = load_workbook(procoreExportFile, data_only=True)
    procore_workbook =map_template["Importer Data Fields"]

    for rowNumber in range(1, procore_workbook.max_row):
        procoreCostCode= str(procore_workbook["A"+str(rowNumber)].value)
        procoreCostDes= str(procore_workbook["B"+str(rowNumber)].value)
        if len(procoreCostCode) == 11:
            ProcoreDataDict [procoreCostCode]= [procoreCostDes]
   
################################LATEST_COST_CODE_File-----------------------------add in the second index like l-LB-M

    ###map with col in the excel file key = col and value =value of the col
    costTypeList ={
                    "3" :"L", 
                    "4":"LB",
                    "5":"M",
                    "6":"MFG",
                    "7":"S",
                    "8":"E",
                    "9":"O",
                    "10":"OH"}
    
    if ProcoreDataDict:
        standardCostCode_template = load_workbook(LATEST_COST_CODE_File, data_only=True)
        standardCostCode_workbook =standardCostCode_template["FINAL VERSION CC"]

        for rowNumber1 in range(2, standardCostCode_workbook.max_row): ##standardCostCode_workbook.max_row
            standardCostCode_procoreCostCode= str(standardCostCode_workbook["A"+str(rowNumber1)].value)
             
            costTypeValue=""
            for colNumbe1 in range(3, 11):
                #####Cost type value reading from stand 
                costTypeColVlaue=  str(standardCostCode_workbook[get_column_letter(colNumbe1)+str(rowNumber1)].value)
                if costTypeColVlaue != "None" and costTypeColVlaue == "X":
                    for key, val in costTypeList.items():
                        if str(colNumbe1) == key:
                            if costTypeValue =="" :
                                costTypeValue =val
                            else:
                                costTypeValue += "-" +val
            try:
                (ProcoreDataDict [standardCostCode_procoreCostCode]).append(costTypeValue)
            except:
                pass
            
            
    #         # if len(standardCostCode_procoreCostCode) == 11:
    #         #     ProcoreDataDict [standardCostCode_procoreCostCode]= [procoreCostDes]
    # if ProcoreDataDict:
    #     print (ProcoreDataDict)
