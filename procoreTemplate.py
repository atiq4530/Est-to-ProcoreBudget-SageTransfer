# ['tkinter', 'openpyxl', 'numpy', 'calendar', 'tk']

import warnings
from ProcoreExportFileDataRead import *
from ProcoreBaseFunction import *
from EstimateDataExtract import *
from ProcoreToSCPC import *

def Procore_Template(projectEstimateFile, altPoValueFile, procoreTemplateFile): #JobNumber, ZipCode, filePath, gpm, contingency, sales
    
    takeOffAndManHOur_Sheet, takeOffAndManHOurFormula_Sheet= "",""
    
    projectName, template, template_Formula="", "", ""
    ###atlValue that multiple with individual atl price
    ###blanded price multiple with total manhour
    atlVlaue, blandedPrice, breakDownDataFormula_Sheet, atlPo_template=0, 0, "", ""
    
    # project estimate--------------------------------------------------
    try:
        template = load_workbook(projectEstimateFile, data_only=True)
        template_Formula = load_workbook(projectEstimateFile, data_only=False)
        
        if template and template_Formula:
            for sheet in  template.sheetnames:
                if "Project Estimate" in sheet:
                    projectEst_workbook = template[sheet]
                    if projectEst_workbook:
                        ###collecting row and col name------------------
                        projectEstRowAndCol(projectEst_workbook)
                        estBreakdownRowAndCol(template["Est Pharma Sys Breakdown"])
                        projectName= projectEst_workbook["B1"].value
                        
                        ####these are the formula that are being used to grap the markup ex 111111*.85 so we need to collect .85 
                        projectEstmateFormula_Sheet=template_Formula[sheet]
                        breakDownDataFormula_Sheet=template_Formula["Est Pharma Sys Breakdown"]
                        221
                        
                        atlVlaue=projectEstimateAtlantaFormula("AES Pharma System - Atlanta PO Value", "CategoryFromProjectESt", "atlValueCol", projectEstmateFormula_Sheet)
                        
                # elif "Open Bookbreak Down"
                if "Takeoff & Mhrs" in sheet:
                    takeOffAndManhourQty(template[sheet])
                    takeOffAndManHOur_Sheet=template[sheet]
                    takeOffAndManHOurFormula_Sheet=template_Formula[sheet]
                    # takeOff=True
            
    except:
        MessageBox("Openpyxl Error", "Openpy xl cannot read project estimate file")

   
    ###atl po value----------------------------------------------
    if altPoValueFile:
        try:
            atlPo_template = load_workbook(altPoValueFile, data_only=True)   
            altWorkbook= atlPo_template["Sheet1"]
            if altWorkbook:
                ATLPOValueRowAndCol(altWorkbook)
        except:
            MessageBox("Worksheet Error!", "May be ATL PO vlaue dont have Sheet1 tab")

    ### procore template----------------------------------------------------------------------
    procore_template, budgetLineItems_Sheet , unMatchSheet, transferSheet ="", "" , "", ""
    
    try:
        procore_template = load_workbook(procoreTemplateFile, data_only=True)
         ##### budget line sheet
        if "Budget Line Items" in procore_template.sheetnames:
            budgetLineItems_Sheet=procore_template["Budget Line Items"]
        
        ### unmatch sheet
        if "Unmatch Data" not in procore_template.sheetnames: 
            procore_template.create_sheet(title="Unmatch Data")
            unMatchSheet = procore_template["Unmatch Data"]
        else:
            del procore_template["Unmatch Data"]
            procore_template.create_sheet(title="Unmatch Data")
            unMatchSheet = procore_template["Unmatch Data"]
            
        ### intacct transfer file
        ####purpose changing cost type to another and make sure that value dont duplicate
        if "transfer file" not in procore_template.sheetnames: 
            procore_template.create_sheet(title="transfer file")
            transferSheet = procore_template["transfer file"]
        else:
            del procore_template["transfer file"]
            procore_template.create_sheet(title="transfer file")
            transferSheet = procore_template["transfer file"]
    except:
            MessageBox("Worksheet Error!", "Executing Error on Procore Budget")
   
    
    ###### call fun to extract data from project est
    extractData=""
    if ProcoreDataDict and template :
        try:
            extractData= projectEstimate(ProcoreDataDict)    
            
            if atlPo_template:
                atlPoData= ATlPoValueCollect(ProcoreDataDict, atlVlaue)
                if atlPoData and extractData:
                    extractData += atlPoData
            
            if takeOffAndManHOur_Sheet:
                #### AES Unit Cost=L from the Est Pharma Sys Breakdown where they use blanded price
                blandedPrice=projectEstimateAtlantaFormula("Labor to Install Cleanroom", "Category", "AES Unit Cost", breakDownDataFormula_Sheet)
                
                takeOffData=takeOffAndManhourFun(takeOffAndManHOur_Sheet, ProcoreDataDict, blandedPrice, takeOffAndManHOurFormula_Sheet)
                if takeOffData and extractData:
                    extractData += takeOffData
        except:
            pass
        
    if extractData and budgetLineItems_Sheet:   
        WriteToExcelFile(extractData, notMatchingData, budgetLineItems_Sheet, unMatchSheet)
        
        ### apply number and format col
        ExcelFormat(budgetLineItems_Sheet)
           
        ###remove duplicate value by add together and make empty A, B 
        duplicateValueAddUp(budgetLineItems_Sheet, "A", "B", "E", "H")       
            
        # # ###remove empty row 
        removeProCoreEmptyValueRow(budgetLineItems_Sheet, "A", "B")
        
        ####once everything is done for procore 
        # then store the budgetLineItems_Sheet in the array so that can be use for scpc
        budgetLineItemSheetArray.append(budgetLineItems_Sheet)
        budgetLineItemSheetArray.append(transferSheet)
        
        ####Procore Save file path location
        saveFilePath= os.path.join(os.path.dirname(procoreTemplateFile) + "\\" + str(today.strftime("%Y-%m-%d") ) +" "+ str(projectName)+ "-procore-budget.xlsx")
        
        ####scpc save file path
        SCPC_File= os.path.dirname(procoreTemplateFile) + "\\" + str(today.strftime("%Y-%m-%d"))
        
        #####transfer file
        SageIntact_TransferFile_GUI() 
        
        try:
            ####procore save file
            procore_template.save(saveFilePath)
            print ("Procore Save file path location: ", saveFilePath)  ###file gonna be saved where you uploaded the procore tem
            # ####scpc file save
            transferFilePath= WriteToTextFile(SCPC_File)
            print ("Done")
            MessageBox("Successful", 
                    "Procore: \n"
                    + saveFilePath +
                    "\n Transfer File: \n"
                    + transferFilePath)
        except:
            print ("Unfortunately! something went wrong. Try Again")
            MessageBox("Wrong!", "Maybe- Same name excel file is open \n Close and run Again")
                  
        
def fileUpload():
    altPoValueFile=""
    ######getting the cost type
    LATEST_COST_CODE_File =r"Q:\Sales\Estimating\Software\Budget\Template-File\Standard CC and Category List v050324 FINAL.xlsx"
    # projectEstimateFile =r"Q:\Sales\Estimating\Software\Budget\Template-File\CHOP Gene Therapy Project Estimate - 2024-05-22.xlsm"
    # procoreTemplateFile =r"Q:\Sales\Estimating\Software\Budget\Template-File\budget(1).xlsx"
    
    # # # ####test file----------------------------------------------------------------
    # procoreTemplateFile=r"C:\Users\matiq\Desktop\without mapping file\budget(1).xlsx" 
    # # #  ###check file when upload file
    # projectEstimateFile=r"C:\Users\matiq\Desktop\without mapping file\Curia Estimate - 5 - 2024-04-29 - full(1).xlsm" 
    # altPoValueFile=r"C:\Users\matiq\Desktop\without mapping file\curia atlanta(1).xlsx"
    # altPoValueFile=r"C:\Users\matiq\Desktop\without mapping file\test atl.xlsx"
    ########## mappingfile =sys.argv[1]
    
    filePathValidationCheck(LATEST_COST_CODE_File, "File From Q drive cannot read \n Check: "+LATEST_COST_CODE_File)
    
    procoreTemplateFile = filedialog.askopenfilename(initialdir=os.environ['USERPROFILE'],
                                            title = "Select Budget Template that Export from Procore",
                                            filetypes = [("Excel files",
                                                            "*.xlsm .xlsx*"),
                                                        ("all files",
                                                            "*.*")])
    print (procoreTemplateFile)
    # filePathValidationCheck(procoreTemplateFile, "Uploaded Procore Template Cannnot Read")
    
    projectEstimateFile = filedialog.askopenfilename(initialdir=os.environ['USERPROFILE'],
                                            title = "Select Project Est File",
                                            filetypes = [("Excel files",
                                                            "*.xlsm*"),
                                                        ("all files",
                                                            "*.*")])
    # filePathValidationCheck(projectEstimateFile, "Uploaded Project Estimate File Cannnot Read")
    print (projectEstimateFile)
    try:
        altPoValueFile = filedialog.askopenfilename(initialdir=os.environ['USERPROFILE'],
                                                title = "Select Atlanta PO File",
                                                filetypes = [("Excel files",
                                                                "*.xlsx*"),
                                                            ("all files",
                                                                "*.*")])
    except:
        pass
    
    
    if procoreTemplateFile and  LATEST_COST_CODE_File:
        procoreTemplateDataRead(procoreTemplateFile, LATEST_COST_CODE_File)

        if projectEstimateFile:
            Procore_Template(projectEstimateFile, altPoValueFile, procoreTemplateFile)
        else:
            MessageBox("Error!", "project est file cannot read")
    else:
        MessageBox("Error!", "Procore Template File or LATEST COST CODE File having issues to read")
        
if __name__ == '__main__':
    import importlib.util
    import subprocess
    
    package = ['tkinter', 'openpyxl', 'numpy', 'calendar', 'tk', 'csv']
    for name in package:
        if (spec := importlib.util.find_spec(name)) is not None:
            print(f"{name!r} has imported")
        else:
            print(f"can't find the {name!r} module")
            subprocess.call(['pip3', 'install', name])
    print ("Runnnig without mapping file....")
    
    fileUpload()
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
      