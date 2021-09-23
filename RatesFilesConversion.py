# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import sys
from os import listdir
    
    
def format_conversion_currPremPerk_g(product, filePath, sheets, firstRow = 4):

    BAND_LIST = ['B1', 'B2', 'B3', 'B4', 'B5']   
    RISK_CLASS_DICT = {'1':'UPNT', '2':'SPNT', '3':'NT', '4':'ST', '5':'T'}
    RISK_CLASS_1_DICT = {'1':'UP', '2':'SP', '3':'', '4':'SP', '5':''}
    RISK_CLASS_2_DICT = {'1':'NT', '2':'NT', '3':'NT', '4':'T', '5':'T'}
    output = pd.DataFrame([])
    df_combine = pd.DataFrame([])
    
    for sheet in sheets:
        
        df = pd.read_excel(filePath, sheet_name=sheet, skiprows=firstRow-1, encoding="utf8")
        gender_and_risk_class = df.columns[1:6].tolist()
        df = pd.melt(df, id_vars = ["Age"], value_vars = gender_and_risk_class)    
        tmp = df["variable"].str.split(pat=r'([A-Za-z]+)', expand=True).iloc[:,1:3]
        if 'Band' in sheet:
            df['Band'] = "B"+ sheet.split(' ')[-1]
        df["Gender"] = tmp[1]
        df["Class"] = tmp[2]
        df["Class_1"] = tmp[2]
        df["Class_2"] = tmp[2]
        df["Table Rating"] = "-"
        df["Product"] = product

        df_combine = pd.concat([df_combine,df])
    
    if len(sheets) == 3:
        for band in BAND_LIST:
            df_band = df_combine.copy(deep = True)
            df_band["Band"] = band
            output = pd.concat([output, df_band])
    else:
        output = df_combine
        
    output = output.rename(columns = {"value": "Premium_Rate", "Age": "Issue Age"})   
    output = output.replace({"Class":RISK_CLASS_DICT, "Class_1":RISK_CLASS_1_DICT, "Class_2":RISK_CLASS_2_DICT}) 
    output["PA_Key"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Band"] + "," + output["Class_1"] + "," + output["Class_2"] + "," + output["Issue Age"].astype(str)
    output["Code"] = output["Product"] + "," + output["Gender"] + "," + output["Band"] + "," + output["Class"] + "," + output["Table Rating"] + "," + output["Issue Age"].astype(str)
    output = output.drop(columns=["Class_1", "Class_2", "variable"])
    output = output.reset_index(drop=True)
    return output
    
    

def format_conversion_currPremPerk_sub(product, filePath, sheets, firstRow = 4):
    
    output = pd.DataFrame([])
    TABLE_RATINGS = ["A","B","C","D","E","F","H","J","L","P"]
    RISK_CLASS_DICT = {'TOB':'T'}
        
    for sheet in sheets:
        df = pd.read_excel(filePath, sheet_name=sheet, skiprows=firstRow-1, encoding="utf8")
        df = pd.melt(df, id_vars = ["Age"], value_vars = TABLE_RATINGS)    
        df["Gender"] = sheet.split("_")[-2][0]
        df["Class"] = sheet.split("_")[-1]
        df["Band"] = ""        
        df["Product"] = product 
        output = pd.concat([output,df])
    output = output.replace({"Class": RISK_CLASS_DICT})
    output = output.rename(columns = {"variable": "Table Rating", "value": "Premium_Rate", "Age": "Issue Age"}) 
    output["PA_Key"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Class"] + "," + output["Table Rating"] + "," + output["Issue Age"].astype(str)
    output["Code"] = output["Product"] + "," + output["Gender"] + "," + output["Band"] + "," + output["Class"] + "," + output["Table Rating"] + "," + output["Issue Age"].astype(str)
    output = output.reset_index(drop=True)
    return output
    

    

def format_conversion_currPremPerk(product, filePath):
    
    SHEETS_SUB = ["Sub_Classified_Prem_Male_NT", "Sub_Classified_Prem_Male_TOB", "Sub_Classified_Prem_Female_NT"
                  , "Sub_Classified_Prem_Female_TOB","Sub_Classified_Prem_Unisex_NT","Sub_Classified_Prem_Unisex_TOB"]
    SHEETS_G_1 = ["Prem Male", "Prem Female", "Prem Unisex"]
    SHEETS_G_2 = ["Prem Male Band 1", "Prem Female Band 1", "Prem Unisex Band 1", "Prem Male Band 2", "Prem Female Band 2", "Prem Unisex Band 2", "Prem Male Band 3", "Prem Female Band 3"
                  , "Prem Unisex Band 3", "Prem Male Band 4", "Prem Female Band 4", "Prem Unisex Band 4", "Prem Male Band 5"
                  , "Prem Female Band 5", "Prem Unisex Band 5"]
    SHEETS_G_3 = ["Prem Male Band 2", "Prem Female Band 2", "Prem Unisex Band 2", "Prem Male Band 3", "Prem Female Band 3"
                  , "Prem Unisex Band 3", "Prem Male Band 4", "Prem Female Band 4", "Prem Unisex Band 4", "Prem Male Band 5"
                  , "Prem Female Band 5", "Prem Unisex Band 5"]


    output = pd.DataFrame([])
    
    xl = pd.ExcelFile(filePath)
    
    # Validate sheet names for general currPrem
    if all(item in xl.sheet_names for item in SHEETS_G_1):
        sheets_g = SHEETS_G_1
    elif all(item in xl.sheet_names for item in SHEETS_G_2):
        sheets_g = SHEETS_G_2
    elif all(item in xl.sheet_names for item in SHEETS_G_3):
        sheets_g = SHEETS_G_3
    else:
        sys.exit("format_conversion_currPremPerk : Please check general currPrem of input file. Program Terminated.")    
    
    # Validate sheet names for sub classified Prem
    if all(item in xl.sheet_names for item in SHEETS_SUB):
        sheets_sub = SHEETS_SUB
    else:
        sys.exit("format_conversion_currPremPerk : Please check sub classified Prem of input file. Program Terminated.")           
             
    df1 = format_conversion_currPremPerk_g(product, filePath, sheets_g)
    df2 = format_conversion_currPremPerk_sub(product, filePath, sheets_sub)
    output = pd.concat([df1, df2])
    
    output = output[['Product', 'Gender', 'Band', 'Class', 'Table Rating', 'Issue Age', 'PA_Key', 'Code', 'Premium_Rate']]
    output = output.reset_index(drop=True)
       
    return output


def format_conversion_waiverPerk(product, filePath, sheets, firstRow):

    BAND_LIST = ['B2', 'B3', 'B4', 'B5']   
    RISK_CLASS_DICT = {'1':'UPNT', '2':'SPNT', '3':'NT', '4':'ST', '5':'T'}
    RISK_CLASS_1_DICT = {'1':'UP', '2':'SP', '3':'', '4':'SP', '5':''}
    RISK_CLASS_2_DICT = {'1':'NT', '2':'NT', '3':'NT', '4':'T', '5':'T'}
    output = pd.DataFrame([])
    df_combine = pd.DataFrame([])
    
    for sheet in sheets:
        
        df = pd.read_excel(filePath, sheet_name=sheet, skiprows=firstRow-1, encoding="utf8")
        gender_and_risk_class = df.columns[1:6].tolist()
        df = pd.melt(df, id_vars = ["Age"], value_vars = gender_and_risk_class)    
        tmp = df["variable"].str.split(pat=r'([A-Za-z]+)', expand=True).iloc[:,1:3]
        if 'Band' in sheet:
            df['Band'] = "B"+ sheet.split(' ')[-1]
        df["Gender"] = tmp[1]
        df["Class"] = tmp[2]
        df["Class_1"] = tmp[2]
        df["Class_2"] = tmp[2]
        df["Product"] = product

        df_combine = pd.concat([df_combine,df])
    
    if len(sheets) == 3:
        for band in BAND_LIST:
            df_band = df_combine.copy(deep = True)
            df_band["Band"] = band
            output = pd.concat([output, df_band])
    else:
        output = df_combine
        
    output = output.rename(columns = {"value": "Premium_Rate", "Age": "Issue Age"})   
    output = output.replace({"Class":RISK_CLASS_DICT, "Class_1":RISK_CLASS_1_DICT, "Class_2":RISK_CLASS_2_DICT}) 
    output["PA_Key"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Band"] + "," + output["Class_1"] + "," + output["Class_2"] + "," + output["Issue Age"].astype(str)
    output["Code"] = output["Product"] + "," + output["Gender"] + "," + output["Band"] + "," + output["Class"] + "," + output["Issue Age"].astype(str)
    output = output.drop(columns=["Class_1", "Class_2", "variable"])
    output = output[['Product', 'Gender', 'Band', 'Class', 'Issue Age', 'PA_Key', 'Code',  'Premium_Rate']]
    output = output.reset_index(drop=True)
    return output

def format_conversion_NSP(product, filePath, firstRow):
    
    output = pd.DataFrame([])
    
    xl = pd.ExcelFile(filePath)

    for sheet in xl.sheet_names:
        df = xl.parse(sheet_name = sheet, skiprows=firstRow-1, encoding="utf8")
        tmp = sheet.split(' ')
        df["Gender"] = tmp[0][0]
        if len(tmp) == 2:
            df["Table Rating"] = "-"
        else:
            df["Table Rating"] = tmp[-1]
        
        df["Product"] = product
        output = pd.concat([output, df])   
    output["PA_Code"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Table Rating"] + "," + output["Age"].astype(str)
    output["CODE"] = output["Product"] + "," + output["Gender"] + "," + output["Table Rating"] + "," + output["Age"].astype(str)
    output.columns = output.columns.astype(str)
    output["0"] = output["1"]
    output["-1"] = output["1"]
    output = output[['Product', 'Gender', 'Table Rating', 'Age', 'PA_Code', 'CODE'] + [ str(i) for i in range(-1,123)]]
    output.columns = ['Product', 'Gender', 'Table Rating', 'Age', 'PA_Code', 'CODE'] + [ 'Dur.' + str(i) for i in range(-2,122)]
    
    return output

def format_conversion_BOYStateReserve(product, filePath, firstRow):
    
    output = pd.DataFrame([])
    
    xl = pd.ExcelFile(filePath)

    for sheet in xl.sheet_names:
        df = xl.parse(sheet_name = sheet, skiprows=firstRow-1, encoding="utf8")
        tmp = sheet.split(' ')
        df["Gender"] = tmp[0][0]
        if len(tmp) == 2:
            df["Table Rating"] = "-"
        else:
            df["Table Rating"] = tmp[-1]
        
        df["Product"] = product
        output = pd.concat([output, df])   
        
    output["PA_Code"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Table Rating"] + "," + output["Age"].astype(str)
    output["CODE"] = output["Product"] + "," + output["Gender"] + "," + output["Table Rating"] + "," + output["Age"].astype(str)
    output.columns = output.columns.astype(str)
    output["0"] = output["1"]
    output["-1"] = output["1"]
    output = output[['Product', 'Gender', 'Table Rating', 'Age', 'PA_Code', 'CODE'] + [ str(i) for i in range(-1,123)]]
    output.columns = ['Product', 'Gender', 'Table Rating', 'Age', 'PA_Code', 'CODE'] + [ 'Dur.' + str(i) for i in range(-2,122)]
    
    return output    


def format_conversion_cashValuePerK(product, filePath, firstRow=5):

    SHEETS_1 = ["Male CV(BOY)", "Female CV(BOY)", "Unisex CV(BOY)"]    
    SHEETS_2 = ["Male UPNT CV (BOY)", "Male SPNT CV (BOY)", "Male NT CV (BOY)", "Male SPT CV (BOY)", "Male SPT CV (BOY)"
          , "Female UPNT CV (BOY)", "Female SPNT CV (BOY)", "Female NT CV (BOY)", "Female SPT CV (BOY)", "Female TOB CV (BOY)"
          , "Unisex UPNT CV (BOY)", "Unisex SPNT CV (BOY)", "Unisex NT CV (BOY)", "Unisex SPT CV (BOY)", "Unisex TOB CV (BOY)"]
    SHEETS_3 = ["GCV (BOY) Male", "GCV (BOY) Female", "GCV (BOY) Unisex"]
    
    
    output = pd.DataFrame([])
    
    xl = pd.ExcelFile(filePath)
    
    if all(item in xl.sheet_names for item in SHEETS_1):
        sheets = SHEETS_1
    elif all(item in xl.sheet_names for item in SHEETS_2):
        sheets = SHEETS_2
    elif all(item in xl.sheet_names for item in SHEETS_3):    
        sheets = SHEETS_3
    else:
        sys.exit("format_conversion_cashValuePerK : Please check cash value of input file. Program Terminated.")
    
    for sheet in sheets:
        df = xl.parse(sheet_name = sheet, skiprows=firstRow-1, encoding="utf8")
        df["Gender"] = get_gender(sheet)
        df["Class"] = get_class(sheet)        
        df["Product"] = product
        output = pd.concat([output, df])           
        
    output["PA_Key"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Class"] + "," + output["Age"].astype(str)
    output["CODE"] = output["Product"] + "," + output["Gender"] + "," + output["Class"] + "," + output["Age"].astype(str)
    output.columns = output.columns.astype(str)
    output = output[['Product', 'Gender', 'Class', 'Age', 'PA_Key', 'CODE'] + [ str(i) for i in range(0,122)]]
    output.columns = ['Product', 'Gender', 'Class', 'Age', 'PA_Key', 'CODE'] + [ 'Dur.' + str(i) for i in range(0,122)]
    
    return output


def get_gender(string):
    
    if "Male" in string:
        return "M"
    elif "Female" in string:
        return "F"
    elif "Unisex" in string:
        return "U"
    # For dividends
    elif "Qual" in string:
        return "U"
    else:
        return ""

def get_class(string):

    if "UPNT" in string:
        return "UPNT"
    elif "SPNT" in string:
        return "SPNT"
    elif "NT" in string:
        return "NT"
    elif "SPT" in string:
        return "SPT"
    elif "TOB" in string:
        return "TOB"
    else:
        return ""


def get_product(string):
    
    if "LP10" in string:
        return "L10"
    elif "LP12" in string:
        return "L12"
    elif "LP15" in string:
        return "L15"
    elif "LP20" in string:
        return "L20"
    elif "LP65" in string:
        return "L65"
    elif "HECV" in string:
        return "L85"
    elif "L100" in string:
        return "L100"
    else:
        return ""
 
def get_dividend_type(string):
    
    string = string.upper()
    if 'DIV' in string:
        return 'Base'
    elif 'PUA' in string:
        return 'PUA'
    elif 'RPU' in string:
        return 'RPU'
    elif 'LISR' in string:
        return 'LISR'
    elif 'ALIR' in string:
        return 'ALIR'
    else:
        return ''
 
def get_dividend_market(string):
    
    if 'Qual' in string:
        return 'Q'
    else:
        return 'NQ'

def get_dividend_risk_class(string):
    
    if "UPNT" in string:
        return "UPNT"
    elif "SPNT" in string:
        return "SPNT"
    elif "NT" in string:
        return "NT"
    elif "SPT" in string:
        return "ST"
    elif "TOB" in string:
        return "T"
    else:
        sys.exit("get_dividend_risk_class : Please check risk class of input file. Program Terminated with value: " + string)  
 
def get_dividend_risk_subclass_1(string):
    
    if 'SP' in string:
        return 'SP'
    elif 'UP' in string:
        return 'UP'
    else:
        return ''

def get_dividend_risk_subclass_2(string):
    
    if "NT" in string :
        return "NT"
    elif 'SPT' in string or 'TOB' in string:
        return 'T'
    else:
        sys.exit("get_dividend_risk_subclass_2 : Please check risk class of input file. Program Terminated with value: " + string)  
    
def format_conversion_TAI_TR(product, filePath, firstRow=4):

    SHEETS_1 = ["Male NS", "Male SM", "Female NS", "Female SM", "Unisex NS", "Unisex SM"]    
    SHEETS_2 = ["TAI TR Male UPNT", "TAI TR Male SPNT", "TAI TR Male NT", "TAI TR Male SPT", "TAI TR Male TOB"
                , "TAI TR Female UPNT", "TAI TR Female SPNT", "TAI TR Female NT", "TAI TR Female SPT", "TAI TR Female TOB"
                , "TAI TR Unisex UPNT", "TAI TR Unisex SPNT", "TAI TR Unisex NT", "TAI TR Unisex SPT", "TAI TR Unisex TOB"]
    
    output = pd.DataFrame([])
    
    xl = pd.ExcelFile(filePath)
    
    if all(item in xl.sheet_names for item in SHEETS_1):
        sheets = SHEETS_1
    elif all(item in xl.sheet_names for item in SHEETS_2):
        sheets = SHEETS_2
    else:
        sys.exit("format_conversion_TAI_TR : Please check sheet name of input file. Program Terminated.")
    
    for sheet in sheets:
        df = xl.parse(sheet_name = sheet, skiprows=firstRow-1, encoding="utf8")
        risk_class = sheet.split(" ")[-1]
        if len(sheets) == 6:
            df["Smk Stat"] = "NT" if "NS" in sheet else "T"
            df["Code_class"] = "N" if "NS" in sheet else "S"      
        else:
            df["Smk Stat"] = "T" if "TOB" in sheet else risk_class
            df["Code_class"] = "NT" if "NS" in sheet else "S"
        df["Product"] = product
        df["Gender"] = get_gender(sheet)
        output = pd.concat([output, df])           
        
    output["PA_Key"] = "CP" + output["Product"] + "A," + output["Gender"] + "," + output["Smk Stat"] + "," + output["Age"].astype(str)
    output["Code"] = output["Product"] + "," + output["Gender"] + "," + output["Code_class"] + "," + output["Age"].astype(str)
    output.columns = output.columns.astype(str)
    output = output[['PA_Key', 'Product', 'Gender', 'Smk Stat', 'Age', 'Code'] + [ str(i) for i in range(1,122)]]
    output.columns = ['PA_Key', 'Product', 'Gender', 'Smk Stat', 'Age', 'Code'] + [ 'Dur.' + str(i) for i in range(1,122)]
    
    return output

def format_conversion_dividends(product, filePath, firstRow=6):
    
    output = pd.DataFrame([])
    xl = pd.ExcelFile(filePath)
    
    for sheet in xl.sheet_names:
        df = xl.parse(sheet_name = sheet, skiprows = firstRow - 1, encoding='utf8')
        info = sheet.split(' ')
        
        # Set Base/PUA/RPU
        df['Base/PUA/RPU'] = get_dividend_type(info[0])
        df['Gender'] = get_gender(info[1])
        df['Market'] = get_dividend_market(info[1])
        df['Market_in_key'] = get_dividend_market(info[1]) if get_gender(info[1]) == 'U' else ''
        df['Underwriting Class'] = get_dividend_risk_class(info[2])
        df['risk_class_in_key_1'] = get_dividend_risk_subclass_1(info[2])
        df['risk_class_in_key_2'] = get_dividend_risk_subclass_2(info[2])
        df['Band'] = info[3] if len(info) == 4 else ''
        output = pd.concat([output, df])
    
    output['Product'] = product
    output['PA_KEY'] = "CP" + output['Product'] + 'A,' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market_in_key'] + ',' + \
                        output['risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
    
    output['CODE'] = output['Product'] + ',' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market'] + ',' + \
                        output['Underwriting Class'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
    output['Dur.0'] = 0
    output.columns = output.columns.astype(str)
    output = output[['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Age', 'PA_KEY', 'CODE', 'Dur.0'] + [str(i) for i in range(1,122)]]
    output.columns = ['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Iss. Age', 'PA_KEY', 'CODE'] + [ 'Dur.' + str(i) for i in range(0,122)]
    
    return output

def validation(srcFile, desFile):

    df1 = pd.read_csv(srcFile).fillna(0)
    df2 = pd.read_csv(desFile).fillna(0)
    test = df1-df2    
    return (test == 0).any().any()

if __name__ == '__main__':
    

    
    ##############################
    ####    ^Dividends    ####
    ##############################       
    
    dividend = pd.DataFrame([])
    inputDir = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\Dividend Rate File 9.20.2021' 
    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\Dividends_v7_2021.xlsx'
     
    fileNameList = listdir(inputDir)
     
    for eachFile in fileNameList:
        product = get_product(eachFile)
        dividend = pd.concat([dividend, format_conversion_dividends(product, inputDir+'\\'+eachFile)])
         
    dividend = dividend.reset_index(drop=True)
    
    with pd.ExcelWriter(outputFilePath, mode='a', engine='openpyxl') as writer:
        dividend.to_excel(writer, sheet_name='^Dividends_v7_2021')
    
    ##############################
    ####    ^CashValuePerK    ####
    ##############################       
    # Need to rename the input file name with Product name as prefix
    # Modify rate file for L10
    
#    cashValue = pd.DataFrame([])
#    currPrem = pd.DataFrame([])
#    tai_tr = pd.DataFrame([])
#    inputDir = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\PremiumRates'
#    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\CashValuePreK.xlsx'
##    inputDir = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\TAI_TR'
##    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\TAI_IR.xlsx'
#    fileNameList = listdir(inputDir)  
#    
#    for eachFile in fileNameList:
#        product = get_product(eachFile)
##        cashValue = pd.concat([cashValue, format_conversion_cashValuePerK(product, inputDir+"\\"+eachFile)])
#        currPrem = pd.concat([currPrem, format_conversion_currPremPerk(product, inputDir+"\\"+eachFile)])
##        tai_tr = pd.concat([tai_tr, format_conversion_TAI_TR(product, inputDir+"\\"+eachFile)])
#    # Reset index
##    cashValue = cashValue.reset_index(drop=True)  
#    currPrem = currPrem.reset_index(drop=True)
#    tai_tr = tai_tr.reset_index(drop=True)
#    
#    # Write to excel file      
#    with pd.ExcelWriter(outputFilePath, mode='a', engine="openpyxl") as writer:
##        cashValue.to_excel(writer, sheet_name='CashValuePreK')   
#        currPrem.to_excel(writer, sheet_name='CurrPremPerK_L100')  
#        tai_tr.to_excel(writer, sheet_name='TAI_TR')
    
    
    
    ##############################
    ####    ^CurrPremPerK     ####
    ##############################   
    
#    product = 'L15'
#    output = pd.DataFrame([])
#    
#    for eachFile in fileNameList:
#        tmp
#    
#    
##    sheet_g = ["Prem Male", "Prem Female", "Prem Unisex"]
#    sheet_g = ["Prem Male Band 2", "Prem Female Band 2", "Prem Unisex Band 2", "Prem Male Band 3", "Prem Female Band 3", "Prem Unisex Band 3", "Prem Male Band 4", "Prem Female Band 4", "Prem Unisex Band 4", "Prem Male Band 5", "Prem Female Band 5", "Prem Unisex Band 5"]
#    firstRow_g = 4
#    firstRow_sub = 4
#    sheet_sub = ["Sub_Classified_Prem_Male_NT", "Sub_Classified_Prem_Male_TOB", "Sub_Classified_Prem_Female_NT", "Sub_Classified_Prem_Female_TOB","Sub_Classified_Prem_Unisex_NT","Sub_Classified_Prem_Unisex_TOB"]
#    
#    filePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\LP15 (2021)\\Laurie W\\LP15 2021 Premium Rate File 7.30.21 - Values.xlsx'
#    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\CurrPremPerK.xlsx'
#   
#    output = format_conversion_currPremPerk(product, filePath, sheet_g, sheet_sub, firstRow_g, firstRow_sub)
#    with pd.ExcelWriter(outputFilePath, mode='a', engine="openpyxl") as writer:
#        output.to_excel(writer, sheet_name='currPremPerk_LP15_new')     
#        
        
    ##############################
    ####     ^WaiverPerK      ####
    ##############################        
        
#    product = 'L15'
##    sheets = ["WP Male", "WP Female", "WP Unisex"]
#    sheets = ["WP Male Band 2", "WP Female Band 2", "WP Unisex Band 2", "WP Male Band 3", "WP Female Band 3", "WP Unisex Band 3", "WP Male Band 4", "WP Female Band 4", "WP Unisex Band 4", "WP Male Band 5", "WP Female Band 5", "WP Unisex Band 5"]
##    firstRow = 6
#    firstRow = 4
##    sheet_sub = ["Sub_Classified_Prem_Male_NT", "Sub_Classified_Prem_Male_TOB", "Sub_Classified_Prem_Female_NT", "Sub_Classified_Prem_Female_TOB","Sub_Classified_Prem_Unisex_NT","Sub_Classified_Prem_Unisex_TOB"]
#    
#    filePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\LP15 (2021)\\Laurie W\\LP15 2021 Premium Rate File 7.30.21 - Values.xlsx'
#    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\WaiverPerk.xlsx'
#   
#    output = format_conversion_waiverPerk(product, filePath, sheets, firstRow)
#    with pd.ExcelWriter(outputFilePath, mode='a', engine="openpyxl") as writer:
#        output.to_excel(writer, sheet_name='waiverPerk_LP15_new')             


    ##############################
    ####         ^NSP         ####
    ############################## 

    
#    product = 'L100'
#    sheets = ["WP Male", "WP Female", "WP Unisex"]
##    sheets = ["WP Male Band 2", "WP Female Band 2", "WP Unisex Band 2", "WP Male Band 3", "WP Female Band 3", "WP Unisex Band 3", "WP Male Band 4", "WP Female Band 4", "WP Unisex Band 4", "WP Male Band 5", "WP Female Band 5", "WP Unisex Band 5"]
##    firstRow = 6
#    firstRow = 6
##    sheet_sub = ["Sub_Classified_Prem_Male_NT", "Sub_Classified_Prem_Male_TOB", "Sub_Classified_Prem_Female_NT", "Sub_Classified_Prem_Female_TOB","Sub_Classified_Prem_Unisex_NT","Sub_Classified_Prem_Unisex_TOB"]
#    filePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\NSP\\CSO17 Standard & Perm Substandard 3.75% NSP Rate File.xlsx'
##    filePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\test.xlsx'
#    outputFilePath = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\NSP.xlsx'
#    
#    xl = pd.ExcelFile(filePath)
#
#    xl.sheet_names  # see all sheet names
#
#    xl.parse(sheet_name = "Male NSP", skiprows=5)  # read a specific sheet to DataFrame
#    
#    output = format_conversion_NSP(product, filePath, firstRow)
#    with pd.ExcelWriter(outputFilePath, mode='a', engine="openpyxl") as writer:
#        output.to_excel(writer, sheet_name='L100')    


    ##############################
    ####      Validation      ####
    ############################## 

    
#    srcFile = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\data1.csv'
#    desFile = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\data2.csv'
#    test = validation(srcFile, desFile)
    