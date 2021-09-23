import numpy as np
import pandas as pd
import sys
from os import listdir
import xlrd
import configparser
from openpyxl import load_workbook



class base():

    def __init__(self, input_file, first_row, output_column_names):
        self.input_file = input_file
        self.first_row = first_row
        self.output_column_names = output_column_names
        self.product_name = self._get_product(self.input_file)

    def _get_product(self, string):

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
            sys.exit("get_product : Please check the name of input file. Program Terminated with value: " + string)

    def _get_gender(self,string):

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

    def _get_risk_class(self, string):
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
            sys.exit("get_dividend_risk_class : Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_risk_subclass_1(self, string):

        if 'SP' in string:
            return 'SP'
        elif 'UP' in string:
            return 'UP'
        else:
            return ''

    def _get_risk_subclass_2(self, string):

        if "NT" in string :
            return "NT"
        elif 'SPT' in string or 'TOB' in string:
            return 'T'
        else:
            sys.exit("get_risk_subclass_2 : Please check the sheet name of input file. Program Terminated with value: " + string)

class dividendParser(base):

    FIRST_ROW_DEFAULT = 6
    SHEET_NAME_DEFAULT = '^Dividends_v7_2021_test'
    COLUMN_NAMES_DEFAULT = ['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Iss. Age', 'PA_KEY', 'CODE'] +\
                   [ 'Dur.' + str(i) for i in range(0,122)]

    def __init__(self, input_file, first_row = FIRST_ROW_DEFAULT, output_column_names = COLUMN_NAMES_DEFAULT):
        self.input_file = input_file
        self.first_row = first_row
        self.output_column_names = output_column_names



    def _get_risk_class(self, string):
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
            sys.exit("get_risk_class : Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_dividend_market(self, string):

        if 'Qual' in string:
            return 'Q'
        else:
            return 'NQ'

    def _get_dividend_type(self, string):

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
            sys.exit("get_dividend_type : Please check the sheet name of input file. Program Terminated with value: " + string)

    def parse(self):

        output = pd.DataFrame([])
        xl = pd.ExcelFile(self.input_file)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet_name = sheet, skiprows = self.first_row - 1, encoding='utf8')
            # Parse info based on sheet name
            info = sheet.split(' ')

            # Set Base/PUA/RPU
            df['Base/PUA/RPU'] = self._get_dividend_type(info[0])
            df['Gender'] = self._get_gender(info[1])
            df['Market'] = self._get_dividend_market(info[1])
            df['Market_in_key'] = self._get_dividend_market(info[1]) if self._get_gender(info[1]) == 'U' else ''
            df['Underwriting Class'] = self._get_risk_class(info[2])
            df['risk_class_in_key_1'] = self._get_risk_subclass_1(info[2])
            df['risk_class_in_key_2'] = self._get_risk_subclass_2(info[2])
            df['Band'] = info[3] if len(info) == 4 else ''
            output = pd.concat([output, df])

        output['Product'] = self._get_product(self.input_file)
        # Build PA_KEY
        output['PA_KEY'] = "CP" + output['Product'] + 'A,' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market_in_key'] + ',' + \
                            output['risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
        # Build CODE
        output['CODE'] = output['Product'] + ',' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market'] + ',' + \
                            output['Underwriting Class'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
        output['Dur.0'] = 0
        output.columns = output.columns.astype(str)
        # Rearrange the column
        output = output[['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Age', 'PA_KEY', 'CODE', 'Dur.0'] + [str(i) for i in range(1,122)]]
        # Rename the column
        output.columns = self.output_column_names
        return output




class currPremPerkParser(base):
    FIRST_ROW_DEFAULT = 4
    SHEET_NAME_DEFAULT = '^CurrPremPerK_v7_2021_test'


    def parse(self):
        output = pd.DataFrame([])
        xl = pd.ExcelFile(self.input_file)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet_name = sheet, skiprows = self.first_row - 1, encoding='utf8')
            # Parse info based on sheet name
            info = sheet.split(' ')

def main():

    # Load configuration file
    config = configparser.ConfigParser()
    config.read('config.txt')

    app_args = dict()
    input_dir_rate_file = ''


    ##############################
    ####    ^Dividends    ####
    ##############################

    input_dir_dividend_file = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\Dividend Test'
    output_file = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\Dividends_v7_2021 9.23.21.xlsx'

    df_output = pd.DataFrame([])
    input_file_list = listdir(input_dir_dividend_file)



    for eachFile in input_file_list:
        parser = dividendParser(input_file = input_dir_dividend_file+'\\'+eachFile)
        df_output = pd.concat([df_output, parser.parse()])

    df_output = df_output.reset_index(drop=True)

    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name=dividendParser.SHEET_NAME_DEFAULT)

if __name__ == '__main__':
    main()