
from ProcoreBaseFunction import *

budgetLineItemSheetArray=[]
intactTransDataArr=[]
JobNumber_arr=[]
Project_Name_arr=[]

# PROJECTID=18-1128
# PJESTIMATEID= E01

def Intact_Transfer_File(ProjectNumber, estNumber): #JobNumber, ZipCode, filePath, gpm, contingency, sales
    workbook=budgetLineItemSheetArray[0]
    transferSheet=budgetLineItemSheetArray[1]
   
    if workbook :
        row_heading = JobCostCodeFun("Cost Code", workbook)
        Cost_Code_Col=get_column_letter(ColumnFunction("Cost Code", row_heading, workbook)) 
        Cost_Type_Col=get_column_letter(ColumnFunction("Cost Type", row_heading, workbook)) 
        Unit_Qty_Col=get_column_letter(ColumnFunction("Unit Qty", row_heading, workbook) )
        UOM_Cost_Col=get_column_letter(ColumnFunction("Unit Cost", row_heading, workbook)) ###unit
        UOM_Col=get_column_letter(ColumnFunction("UOM", row_heading, workbook)) ###unit
        
        Budget_Amount_Col=get_column_letter(ColumnFunction("Budget Amount", row_heading, workbook)) 
        activeEst = "Estimate Conversion active"
        active ="active"
        ori ="Original"
        for RowNumber in range(row_heading+1 , workbook.max_row+1):
            RowNumber_str= str(RowNumber)
           
            COST_CODE_Value=str(workbook[Cost_Code_Col+RowNumber_str].value)
            
            Cost_Type_Value= workbook[Cost_Type_Col+RowNumber_str].value       
            #####Cost type
            if Cost_Type_Value == "L":
                  Cost_Type_Value ="LAB"
                  
            elif Cost_Type_Value == "M":
                Cost_Type_Value ="MTL"
                
            elif Cost_Type_Value == "O":
                Cost_Type_Value ="OTH"
                
            elif Cost_Type_Value == "E":
                Cost_Type_Value ="EQU"
                
            elif Cost_Type_Value == "S":
                Cost_Type_Value ="SUB"
            
            elif Cost_Type_Value == "OH":
                Cost_Type_Value ="OTH"
            
            Unit_Qty_Value=workbook[Unit_Qty_Col+RowNumber_str].value
            
            ####UOM
            UOM_Value=workbook[UOM_Col+RowNumber_str].value
            # 
            if UOM_Value == "hours":
                UOM_Value = "HRS"
            UOM_Value = (UOM_Value).upper()
                   
            UOM_Cost_Value=workbook[UOM_Cost_Col+RowNumber_str].value
            
            Budget_Amount_Value=workbook[Budget_Amount_Col+RowNumber_str].value
            try:
                totalBudget = format(float(Budget_Amount_Value), ".2f")
            except:
                pass
            
            transferSheet[Cost_Code_Col+RowNumber_str].value = COST_CODE_Value
            transferSheet[Cost_Type_Col+RowNumber_str].value = Cost_Type_Value
            transferSheet[Unit_Qty_Col+RowNumber_str].value = Unit_Qty_Value
            transferSheet[UOM_Cost_Col+RowNumber_str].value = UOM_Value
            transferSheet[UOM_Col+RowNumber_str].value = UOM_Cost_Value
            transferSheet[Budget_Amount_Col+RowNumber_str].value = totalBudget
            
    
    if transferSheet:
        todayDate = mm+"/"+dd+"/"+yy
        counter =1
        ####add duplicate value 
        duplicateValueAddUp(transferSheet, "A", "B", "E", "H")
        
        removeProCoreEmptyValueRow(transferSheet, "A", "B")
        
        intactTransDataArr.append(["PJESTIMATEID","DESCRIPTION", "STATUS", "ESTIMATEDATE", "PJESTIMATETYPENAME", "PROJECTID",  
                                  "PJESTIMATEENTRY_LINENO", "PJESTIMATEENTRY_TASKID", "PJESTIMATEENTRY_COSTTYPEID",
                                "PJESTIMATEENTRY_EUOM", "PJESTIMATEENTRY_QTY", "PJESTIMATEENTRY_UNITCOST", "PJESTIMATEENTRY_AMOUNT"])
        
        for RowNumber2 in range(1 , transferSheet.max_row+1):
            RowNumber2_str =str(RowNumber2)
            
            Tr_COST_CODE_Value = str(transferSheet[Cost_Code_Col+RowNumber2_str].value)
            Tr_Cost_Type_Value=transferSheet[Cost_Type_Col+RowNumber2_str].value 
            Tr_UOM_Value= transferSheet[UOM_Cost_Col+RowNumber2_str].value
            Tr_Unit_Qty_Value =transferSheet[Unit_Qty_Col+RowNumber2_str].value 
            Tr_UOM_Cost_Value= transferSheet[UOM_Col+RowNumber2_str].value
            Tr_totalBudget= transferSheet[Budget_Amount_Col+RowNumber2_str].value
            
            if Tr_COST_CODE_Value != "None":
                data= [str(ProjectNumber+" "+estNumber),activeEst, active, todayDate, ori, ProjectNumber, counter, Tr_COST_CODE_Value, Tr_Cost_Type_Value, Tr_UOM_Value, Tr_Unit_Qty_Value, Tr_UOM_Cost_Value, Tr_totalBudget]
                
                intactTransDataArr.append(data)
                ProjectNumber,estNumber,activeEst, active, todayDate, ori = "", "", "", "", "", "",      
                counter +=1

                
                

def SageIntact_TransferFile_GUI():
    print ("Sage Intacct_TransferFile_GUI is running...........")
    root =Tk()
    root.withdraw()
    
    def SCPC(projectId, root, pjEstId): 
        pNumberId= projectId.get()
        pjEstNumberId= pjEstId.get()

        if len(pNumberId)>1 and (len(pjEstNumberId)>1):
            root.destroy()
            JobNumber_arr.append(pNumberId)
            Project_Name_arr.append(pjEstNumberId)
            Intact_Transfer_File(pNumberId, pjEstNumberId)
            
        else:
            MessageBox("Invalid!", "Project ID & Estimate id cannot empty")

       
    # zipCode=StringVar(root)
    projectId =StringVar(root)
    pjEstId=StringVar(root)
    
    root.geometry("580x320")
    root.title("SCPC Transfer File")
    root.deiconify() 

    # project Number
    # PROJECTID=projectNumber
    Label(root, text="PROJECT ID *18-1128", font=fontStyle,).pack(anchor = W) 
    Entry(root, textvariable=projectId, width=20,font=fontStyle ).pack(anchor = W)
    
    # PJESTIMATEID=projectName
    Label(root, text="PJ ESTIMATE ID *E01", font=fontStyle, ).pack( anchor = W )
    Entry(root, textvariable=pjEstId, width=20, font=fontStyle).pack( anchor = W )

    # Zip code
    # Label(root, text="Zip Code", font=fontStyle, ).pack( anchor = W )
    # Entry(root, textvariable=zipCode, width=20, font=fontStyle).pack( anchor = W )
    
    Button(root, text = "Intacct Transfer File",command=lambda a=projectId, c=root, d=pjEstId, :SCPC(a, c, d), width=20, font=fontStyle).pack( side = LEFT, padx=10)
    Button(root, text = "Exit", command =lambda: root.destroy(), width=10, font=fontStyle).pack( side = LEFT, padx=10)
    root.mainloop()

def WriteToTextFile(SCPC_File):
    SCPC_File = os.path.join(SCPC_File+ " " + str(JobNumber_arr[0]) + " " + str(Project_Name_arr[0]) + " Transfer File.csv")
    
    with open(SCPC_File,'w',newline='') as csvfile:
        csvWriter = csv.writer(csvfile,delimiter=',')
        csvWriter.writerows(intactTransDataArr)
        csvfile.close()
        return SCPC_File
    


