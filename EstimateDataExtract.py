from ProcoreBaseFunction import *
# from dataReadFromProcoreExportFile import *
notMatchingData=[]

####est haeding cost_type =L 
headingList=[
             "AES Labor to Install Cleanroom Architectural Envelope", 
             "HVAC Construction Services", 
             "HVAC Control Systems", 
             "Mechanical Testing and Balancing", 
             "Electrical Construction Services",
             "Sprinkler Systems", 
             "AES Project Management and Commissioning Services", 
             "AES Engineering and Pre-Construction Management Services"
            ]

estBreakdownRowAndColDict={}
def estBreakdownRowAndCol(workbook):
    row_heading = JobCostCodeFun("Job Cost Code", workbook)
    AES_Unit_Cost_Col=get_column_letter(ColumnFunction("AES Unit Cost", row_heading, workbook))
    Cat_Col=get_column_letter(ColumnFunction("Category", row_heading, workbook))
    
    estBreakdownRowAndColDict[0] =AES_Unit_Cost_Col
    estBreakdownRowAndColDict[1] =Cat_Col

atlRowAndColDict={}
def ATLPOValueRowAndCol(workbook):
    # Description	Lead Time (wk)	Est. Qnty	Unit Cost	Extended Cost
    row_heading = JobCostCodeFun("Line #", workbook)
        
    Cost_Code_Col=get_column_letter(ColumnFunction("Line #", row_heading, workbook)) 
    Cat_Col=get_column_letter(ColumnFunction("Item Code:", row_heading, workbook)) 
    Des_Col=get_column_letter(ColumnFunction("Description", row_heading, workbook)) 
    qty_Col=get_column_letter(ColumnFunction("Est. Qnty", row_heading, workbook)) 
    unitCostType_Col=get_column_letter(ColumnFunction("Unit Cost", row_heading, workbook)) 
    ExtendedCost_Col=get_column_letter(ColumnFunction("Extended Cost", row_heading, workbook)) 
    TotalCost_Col=get_column_letter(ColumnFunction("Extended Cost", row_heading, workbook)+1) 
    
    atlRowAndColDict[0]=row_heading
    atlRowAndColDict[1]=Cost_Code_Col
    atlRowAndColDict[2]=Cat_Col
       
    atlRowAndColDict[3]=Des_Col
    atlRowAndColDict[4]=qty_Col
    atlRowAndColDict[5]=unitCostType_Col
    atlRowAndColDict[6]=ExtendedCost_Col
    
    atlRowAndColDict[7]=workbook
    atlRowAndColDict[8]=TotalCost_Col
    
#####project estimate tab value collect
def ATlPoValueCollect(ProcoreDataDict, poVlaue):
    print (poVlaue, type(poVlaue),"PO value-----------")
    data=[]
    matchRowNumber=[]
    manualCalculation="True"
    if atlRowAndColDict:
        ATLWorkSheet=atlRowAndColDict[7]  
        
        costTypeDict={
            "LAB":"L",
            "MAT":"M",
            "Door Hardware (Standard)":"L",
            "Door Hardware (Custom)":"MFG",
            "Window Componets":"S",
            "Panel Repair Kit":"M",
            "AW-":"",
            "FRT":"O",
            "CRT":"O",
            "_TARIFFS":"O"
        }
        
        ####ATL Material Contingency-------------------------------------------------------            # 
       ##### dictonary that holds key= col from the atl and value=from procore template 
       ####cost type included fixed
        materialHeading={
           "ATL Material Contingency":["99-9999-820", "Project Contingency", "OH"],
           "Factory Office Mgmt Recovery":["99-9999-002","Consumables", "OH"],
           "Factory Floor Mgmt Recovery":["99-9999-002","Consumables", "OH"],
           "Factory Floor Consumables":["99-9999-002", "Consumables", "OH"],
           "ATL Tariff Funding":["99-9999-990","Sales Tax, If applicable", "OH"],

           "AES LiteBeam 04 FT - Overhead":["98-9999-851","Corporate OH", "OH"],
           "AES LiteBeam 05 FT - Overhead":["98-9999-851","Corporate OH", "OH"],
           "AES LiteBeam 06 FT - Overhead":["98-9999-851","Corporate OH", "OH"],
           "AES LiteBeam 08 FT - Overehead":["98-9999-851","Corporate OH","OH"],
           "AES LiteBeam 10 FT - Overhead":["98-9999-851","Corporate OH", "OH"],

           "ATL Design & Mgmt Overhead":["98-9999-852", "Manufacturing OH", "MFG"],
           "ATL Material Overhead":["98-9999-852", "Manufacturing OH", "MFG"],
           
           "ATL Crating Overhead":["98-9999-852","Manufacturing OH","MFG"],
           "ATL Freight Overhead":["98-9999-852","Manufacturing OH","MFG"],
           "ATL Repair Kit Overhead":["98-9999-852","Manufacturing OH", "MFG"]
        }
        #### reading all the materialHeading 
        for rowNumber3 in range(2, atlRowAndColDict[0]+1):
            colA =str(ATLWorkSheet[atlRowAndColDict[1]+str(rowNumber3)].value)
            price =ATLWorkSheet["D"+str(rowNumber3)].value
            
            if "Long Lead Item Summary" in colA:
                break
            
            if colA !="None" and (isinstance(price, (int, float))):
                for keyMat, valMat in materialHeading.items():
                    if keyMat ==  colA:
                        try:
                            price =float(price)
                            if price > 0.0:
                                tPrice = poVlaue*price
                                data.append([valMat[0], valMat[2], valMat[1], "True", '', '','', tPrice]) 
                                matchRowNumber.append(rowNumber3)
                        except:
                            pass

        ####key=line # col from alt po value, value = procore description 
        ###if wrong works here        
        mapAtlLineWithProcore= {"Walls":"Walls","Ceilings":"Ceiling", 
                                "Doors":"Door/Frame", "Windows":"Flush Windows", 
                                "Pharma":"Misc Pharma Materials", 
                                "Trims":"Radius Coving & Trims", "Crates":"Freight", "Freight":"Freight", 
                                "Material - LiteBeam":"Litebeams",
                                
                                ####need to go over---------------------------------------------- 
                                "Door Hardware (Standard)":"Door Hardware", 
                                "Door Hardware (Custom)":"Door Hardware", 
                                
                                "Window Componets":"Windows", 
                                "Airwall/Airwall Brackets":"Airwall / Airwall Brackets", 
                                "Misc. Materials":"Misc Pharma Materials", "Panel Repair Kit":"Panel Repair Kits", 
                                "Consumable Usage":"Consumables",  "Specialty Door Support Frames":"Door/Frame",
                                "Tariffs":"Sales Tax, If applicable", "Shipment Loading Labor":"Freight",
                                }  
        # ###Longest Lead Time ---------------------------------------------------------------------------
        lineCodeHeading, cost_Type ="", ""

        for rowNumber in range(atlRowAndColDict[0]+1, ATLWorkSheet.max_row): 
            lineCodeValue =str(ATLWorkSheet[atlRowAndColDict[1]+str(rowNumber)].value)
            CatValue =str(ATLWorkSheet[atlRowAndColDict[2]+str(rowNumber)].value)

            desValue =str(ATLWorkSheet[atlRowAndColDict[3]+str(rowNumber)].value)
            qtyValue =ATLWorkSheet[atlRowAndColDict[4]+str(rowNumber)].value
            unitCostTypeValue =str(ATLWorkSheet[atlRowAndColDict[5]+str(rowNumber)].value)
            ExtendedCostValue =str(ATLWorkSheet[atlRowAndColDict[6]+str(rowNumber)].value)
            
            TotalCostValue =ATLWorkSheet[atlRowAndColDict[8]+str(rowNumber)].value
            
            ###getting line number Code Code: 1.101 Labor - Walls, 
            if "Code Code:" in lineCodeValue and "Subtotal:" not in lineCodeValue and len(lineCodeValue)>7:
                lineCodeHeading=lineCodeValue
            
            ####Cost type-=---------------------------------
            if CatValue != "None":
                for costTypeKey, costTypeVal in costTypeDict.items():
                    if costTypeKey in CatValue:
                        cost_Type = costTypeVal
                        
                    elif costTypeKey in  lineCodeHeading:
                        cost_Type=costTypeVal
            
            if desValue != "None" and ExtendedCostValue != "None" and TotalCostValue :
                ####extend cost value
                exValue =0
                exMatchValue=re.findall(r"\d{1,4}[.]\d{2}", ExtendedCostValue)
                if exMatchValue:
                    exValue= exMatchValue[0]
                
                ####extend cost value
                try:
                    TotalCostValue =float(TotalCostValue)
                    price = poVlaue*TotalCostValue
                    
                    for mappingKey, mappingValue in  mapAtlLineWithProcore.items():
                        ####Code Code: 1.101 Labor - Walls, taking last portation and map with mapAtlLineWithProcore
                        if lineCodeHeading.endswith(mappingKey): 
                            for proCoreKey, procoreValue in ProcoreDataDict.items(): 
                                if procoreValue[0] ==  mappingValue:
                                    ##### quantity value-----------------
                                    if isinstance (qtyValue , (int, float)):
                                        qtyValue = float(qtyValue)
                                                                          
                                    #########cost type
                                    if len(procoreValue[1]) == 1:
                                        cost_Type= procoreValue[1]
                                    ###when nothing works
                                    if cost_Type =="":
                                        cost_Type ="M"
                                    #####unit of measure
                                    if "hr" in unitCostTypeValue:
                                        unitCostTypeValue ="hours"
                                    elif "mi" in unitCostTypeValue:
                                        unitCostTypeValue =""                                            
                                        
                                    data.append([proCoreKey, cost_Type, procoreValue[0], manualCalculation, qtyValue, unitCostTypeValue, exValue, price]) 
                                    matchRowNumber.append(rowNumber)
                                    # print ([proCoreKey,cost_Type, procoreValue[0], qtyValue, unitCostTypeValue, ExtendedCostValue, TotalCostValue, "atl", rowNumber])
                except:
                    pass
        
        #### unmatch data-----------------------------------------------------
        if matchRowNumber:
            notMatchingData.append(["ATL PO Value"])
            # print(matchRowNumber, "-----------")
            lineCode=""
            for rowNumber7 in range(atlRowAndColDict[0]+1, ATLWorkSheet.max_row):
                lineCodeMatValue =str(ATLWorkSheet[atlRowAndColDict[1]+str(rowNumber7)].value)
                if "Code Code:" in lineCodeMatValue and "Subtotal:" not in lineCodeMatValue and len(lineCodeMatValue)>7:
                    lineCode=lineCodeMatValue
                
                desMatValue =str(ATLWorkSheet[atlRowAndColDict[3]+str(rowNumber7)].value)
                ExtendedCostMatValue =str(ATLWorkSheet[atlRowAndColDict[6]+str(rowNumber7)].value)
                TotalCostMatValue =ATLWorkSheet[atlRowAndColDict[8]+str(rowNumber7)].value
                
                if desMatValue != "None" and ExtendedCostMatValue != "None" and  TotalCostMatValue:
                    if rowNumber7 not in matchRowNumber:
                        notMatchingData.append([rowNumber7, lineCode, desMatValue, TotalCostMatValue])
                        print ([rowNumber7, lineCode, desMatValue, TotalCostMatValue, "atl", rowNumber7])
                        
    return(data)

projectEstRowAndColDict={}
def projectEstRowAndCol(projectEst_workbook):
    row_heading = JobCostCodeFun("Job Cost Code", projectEst_workbook)
    Cost_Code_Col=get_column_letter(ColumnFunction("Job Cost Code", row_heading, projectEst_workbook)) 
    Cat_Col=get_column_letter(ColumnFunction("Category", row_heading, projectEst_workbook)) 
    Qty_Col=get_column_letter(ColumnFunction("Qty", row_heading, projectEst_workbook) )
    Unit_Col=get_column_letter(ColumnFunction("Unit", row_heading, projectEst_workbook) )
    Total_Cost_Col=get_column_letter(ColumnFunction("Total", row_heading, projectEst_workbook)) 
    Value_Col=get_column_letter(ColumnFunction("Value", row_heading, projectEst_workbook)) 
    Tax_Col=get_column_letter(ColumnFunction("Tax", row_heading, projectEst_workbook)) 
    
    projectEstRowAndColDict[0]=row_heading
    projectEstRowAndColDict[1]=Cost_Code_Col
    projectEstRowAndColDict[2]=Cat_Col
    projectEstRowAndColDict[3]=Qty_Col
    projectEstRowAndColDict[4]=Unit_Col
    projectEstRowAndColDict[5]=Total_Cost_Col
    projectEstRowAndColDict[6]=projectEst_workbook
    projectEstRowAndColDict[7]=Value_Col
    projectEstRowAndColDict[8]=Tax_Col
    
def projectEstimate(ProcoreDataDict):
    data=[]
    notMatchDatas=[]
    Cost_Type=""
    endRowNumber=0
    atlTaxValue=0
    heading=""
    manualCalculation="True"
    
    if ProcoreDataDict and projectEstRowAndColDict:
        projectEstimateSheet=projectEstRowAndColDict[6] 
        
        for rowNumber in range(projectEstRowAndColDict[0]+1, projectEstimateSheet.max_row):
            projEstCostCodeValue=str(projectEstimateSheet[projectEstRowAndColDict[1]+str(rowNumber)].value).strip()
            projEstDesValue= str(projectEstimateSheet[projectEstRowAndColDict[2]+str(rowNumber)].value).strip()
            
            for proCoreKey, procoreValue in ProcoreDataDict.items(): ###procore key and values
              
                if projEstCostCodeValue == "99-99-99.001.000":
                    endRowNumber=rowNumber
                    break
                
                if "- Atlanta PO Value" in projEstDesValue:
                    if isinstance(projectEstimateSheet[projectEstRowAndColDict[8]+str(rowNumber)].value, (int, float)):
                        atlTaxValue= float(projectEstimateSheet[projectEstRowAndColDict[8]+str(rowNumber)].value)
                    
                if "Heading" == projEstCostCodeValue:
                    heading =projEstDesValue
                    
                if "Architectural Construction Services" in projEstDesValue:
                    Cost_Type ="S"
                    
                if projEstCostCodeValue != "None" and proCoreKey == projEstCostCodeValue : 
                    totalCost= projectEstimateSheet[projectEstRowAndColDict[5]+str(rowNumber)].value
                    quantityValue= projectEstimateSheet[projectEstRowAndColDict[3]+str(rowNumber)].value
                    unit= str(projectEstimateSheet[projectEstRowAndColDict[4]+str(rowNumber)].value)
                    
                    ######quantity----------------
                    if isinstance(quantityValue, (int, float)): 
                        quantityValue = float(quantityValue)
                        
                    #####unit ------------------
                    if "hr" in unit:
                        unit ="hours"
                    elif "ea" in unit:
                        unit ="ea"
                    if unit == "None":
                        unit =""
                    
                    #####Cost type------------
                    if len(procoreValue[1]) == 1:
                        Cost_Type= procoreValue[1]
                    else:
                        if  heading in headingList:   
                            Cost_Type ="L"
                        else :
                            Cost_Type ="M"   
                    #####total cost---------------------------------------
                    if isinstance(totalCost, (int, float)) and totalCost > 0:
                        data.append([proCoreKey, Cost_Type, procoreValue[0], manualCalculation, quantityValue, unit, "" , totalCost])
                        notMatchDatas.append(rowNumber)  
                                                        
        if atlTaxValue>0:
            data.append(["99-9999-990","T","Sales Tax, If applicable", "True", "", "", "", atlTaxValue])
                        
        # print (data)
        if notMatchDatas:
            notMatchingData.append(["Project Estimate"])
            for rowNumber7 in range(projectEstRowAndColDict[0]+1, endRowNumber-1):
                if rowNumber7 not in notMatchDatas:
                    # "13-21-13.001.000", "99-99-99.001.000", "99-99-99.801"= alt po value, corporate margine, actual sell value
                    if str(projectEstimateSheet[projectEstRowAndColDict[1]+str(rowNumber7)].value).strip() not in ["Heading", "AES Markups", "Project Cost Subtotal:", "Project Cost Total:"]    and \
                        str(projectEstimateSheet[projectEstRowAndColDict[2]+str(rowNumber7)].value).strip() not in ["AES Pharma System - Atlanta PO Value", "Labor to Install Cleanroom",  "AES Labor to Install Cleanroom Architectural Envelope", "Sales Burden", "AES Corporate Margin"] :
                       
                       notMatchingData.append([rowNumber7, projectEstimateSheet["A"+str(rowNumber7)].value, projectEstimateSheet["B"+str(rowNumber7)].value, projectEstimateSheet[projectEstRowAndColDict[5]+str(rowNumber7)].value])
                       
                       print ([rowNumber7, projectEstimateSheet["A"+str(rowNumber7)].value, projectEstimateSheet["B"+str(rowNumber7)].value, projectEstimateSheet[projectEstRowAndColDict[5]+str(rowNumber7)].value, "est"])
        return(data)
                            
takeOffAndManHourQtyData={}

###collecting data up to AES Litebeam rest of them are ea and some of are sf but sf and qty are the same 
def takeOffAndManhourQty(takeOffManHourSheet):
    for  index in range (1, takeOffManHourSheet.max_row):
            if takeOffManHourSheet["B"+str(index)].value is not None :
                description = str(takeOffManHourSheet["B"+str(index)].value)
                qty = takeOffManHourSheet["J"+str(index)].value
                # print( type(qty))
                if "AES Litebeam" == description:
                    break
                
                if description and (isinstance(qty, int) and qty>0) :
                    takeOffAndManHourQtyData[description]= qty

def takeOffAndManhourFun(takeOffManHourSheet, ProcoreDataDict, unitCost, takeOffAndManHOurFormula_Sheet):
    print (unitCost, type(unitCost),"Manhour blended Price-----------")
    ###man hour == unit qty
    
    data=[]
    matchRowNumber=[]
    manualCalculation="True"
    UOM ="hours"
    descriptionHeading ={
                    "DOORS":"L",
                    # "Equipment Rental":"E",
                    "Travel and Expenses":"O",
                    "Overhead":"O"
                    }
    
    desHeading_list =list(descriptionHeading.keys())
    costType_hold=""
    Cost_Type="L"
    
    for  index in range (1, takeOffManHourSheet.max_row):
        costCode = str(takeOffManHourSheet["A"+str(index)].value) ####b= cost code
        description = str(takeOffManHourSheet["B"+str(index)].value) ####b= task
        manHour = str(takeOffManHourSheet["r"+str(index)].value) ####r= total Hour
        manHourFormula = str(takeOffAndManHOurFormula_Sheet["r"+str(index)].value) ####r= total Hour
       
        if description in desHeading_list:
                costType_hold= descriptionHeading[description]
                
        for procoreKey, procoreValue in ProcoreDataDict.items(): 
            if "=SUBTOTAL" not in manHourFormula and costCode != "None": ####not taking if subtotal in manhour 
                if procoreKey == costCode:
                    # print (costCode) 
                        
                    #####this is for Equipment Rental----------------------------      
                    if costCode in ["01-5213-004","01-5213-006","01-5213-005","01-5213-001","01-5213-003"]:
                        pass
                        # manHour= takeOffManHourSheet["O"+str(index)].value
                        # UOM = "weeks" if manHour else ""
                        # Cost_Type ="E"    
                        # unitCostManual =""
                        # total=  takeOffManHourSheet["Z"+str(index)].value
                        # if isinstance(total, (int, float)):
                        #     if total >0:
                        #         data.append([procoreKey, Cost_Type, procoreValue[0], manualCalculation, manHour, UOM, unitCostManual, total])
                    #####this is for Travel and Expenses + Overhead----------------------------      
                    elif costCode in ["01-3200-904","01-3200-906","01-3200-905","01-3200-902","01-3200-901", "01-5213-007", "01-0010-003"]:
                        pass
                        # manHour= takeOffManHourSheet["L"+str(index)].value
                        # UOM = "weeks" if manHour else ""
                        # Cost_Type ="O"    
                        # unitCostManual =""
                        # total=  takeOffManHourSheet["Z"+str(index)].value
                        # if isinstance(total, (int, float)):
                        #     if total >0:
                        #         data.append([procoreKey, Cost_Type, procoreValue[0], manualCalculation, manHour, UOM , unitCostManual, total])
                      
                    else:
                        total=0     
                        try:
                            manHour =int(manHour)
                            if manHour >0:
                                total =float(manHour)*float(unitCost)
                                ######Cost Type---------------------------------------------------
                                ####if one only one type of cost then take it 
                                if len(procoreValue[1]) == 1:
                                    Cost_Type= procoreValue[1]
                                else:
                                    if costType_hold:
                                        Cost_Type =costType_hold                                    
                                    else:
                                        Cost_Type= procoreValue[1][0]
                                        
                                ####exception                                      
                                data.append ([procoreKey, Cost_Type, procoreValue[0], manualCalculation, manHour, UOM, unitCost, total])
                                matchRowNumber.append(index)
                        except:
                            pass
                                
                  
    if matchRowNumber:
        exceList=["None", "Mhrs", "0" , "Man-Hours"]
        notMatchingData.append(["Labor & Take off"])
        for  rowNumber7 in range (1, takeOffManHourSheet.max_row):
            manHour = str(takeOffManHourSheet["r"+str(rowNumber7)].value) ####r= total Hour
            manHour_formula = str(takeOffAndManHOurFormula_Sheet["r"+str(rowNumber7)].value) ####r= total Hour
            # if isinstance(manHour, (int, float)):
            if "=SUBTOTAL" not in manHour_formula and manHour not in exceList:
                if rowNumber7 not in matchRowNumber:
                    notMatchingData.append([rowNumber7, str(takeOffManHourSheet["B"+str(rowNumber7)].value), str(takeOffAndManHOurFormula_Sheet["r"+str(rowNumber7)].value), manHour])
                    print ([rowNumber7, str(takeOffManHourSheet["B"+str(rowNumber7)].value), str(takeOffAndManHOurFormula_Sheet["r"+str(rowNumber7)].value), manHour, "manhour"])
         
    return data

def projectEstimateAtlantaFormula(searchValue, colName, returnColName, sheet):
    returnColNumber= projectEstRowAndColDict[7] if returnColName =="atlValueCol" else estBreakdownRowAndColDict[0]
    colNumber= projectEstRowAndColDict[2] if colName =="CategoryFromProjectESt" else estBreakdownRowAndColDict[1]
        
    result, needValue, currentValue ="", False, ""    ###needValue meaning it has the * in the atlValueCol but cannot extract
    for  rowNumber in range (1, sheet.max_row):
        catValue = str(sheet[colNumber+str(rowNumber)].value)
        
        
        if searchValue in catValue: 
            # ='Takeoff & Mhrs (2)'!R177*64.85    =774632*0.85
            colValue =str(sheet[returnColNumber+str(rowNumber)].value)
            currentValue=colValue
            
            if "*" in colValue:
                ####assuming it has something multiple
                if returnColName =="atlValueCol":
                    needValue =True
                    
                value_sp = colValue.split("*")
                value= (value_sp[len(value_sp)-1])
                
                try :
                    matchVlaue=""
                    if "." in value:
                        matchVlaue= re.findall(r"\d{0,3}[.]\d{2}", value)
                    else:
                        matchVlaue= re.findall(r"^\d{1,3}", value)

                    if matchVlaue:
                        result = matchVlaue[0]
                except:   
                   pass
               
    #### result "" meaning coldnot find blended price or the atl markup price 
    #### in this case take input from user
    if result == "":
        if returnColName =="atlValueCol" :
            if needValue:
            ##################title, message 
                inputFromUser("Atlanta PO Value", "If Atlanta PO Value Discount Provided Please Enter Here: .85")
                if userInputArray:
                    result = userInputArray[0]
            else:
                result =1
                
        if returnColName =="AES Unit Cost":
            ##################title, message 
            inputFromUser("Blended Price",  "Enter Blended Price for Total Manhour Ex: 64.85 \n Current Formula: "+currentValue +"\n We have total Mhrs(R) from Takeoff & Mhrs\n")
            if userInputArray:
                result = userInputArray[0]
    try:
        if result:
            result = float(result)
             ###if nothign is work then 1
        elif result == "":
            result = int(1)
    except:
       pass
        
    # print (result, type(result), "--------------")
    return result  
