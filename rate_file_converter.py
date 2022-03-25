import pandas as pd
import sys
import numpy as np
from os import listdir
import configparser


class BaseParser():
    """
    Base parser class
    """

    def __init__(self, input_file, product_name, first_row):
        """
        :param input_file: Source file directory
        :param product_name: Product name (L10, L15, L20, L65, L85, L100)
        :param first_row: First row with data including headers
        """
        self.input_file = input_file
        self.first_row = first_row
        self.product_name = product_name

    def _get_gender(self, string):
        """
        Utility method to get gender based on input string
        """

        gender_list = {
            'Male': 'M',
            'Female': 'F',
            'Unisex': 'U',
            # For dividends Market Type 'Qualified'
            'Qual': 'U'
        }

        for gender in gender_list:
            if gender in string:
                return gender_list[gender]

        raise ValueError(
            "get_gender: Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_risk_class(self, string):
        """
        Utility method to get risk class based on input string
        """

        class_list = {
            'UPNT': 'UPNT',
            'SPNT': 'SPNT',
            'NT': 'NT',
            'SPT': 'ST',
            'TOB': 'T',
            'T': 'T'
        }

        string = string.upper()

        for type in class_list:
            if type in string:
                return class_list[type]

        raise ValueError(
            "get_risk_class: Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_risk_subclass_1(self, string):
        """
        Utility class to get risk subclass used for PA_Key
        """

        if 'SP' in string:
            return 'SP'
        elif 'UP' in string:
            return 'UP'
        else:
            return ''

    def _get_risk_subclass_2(self, string):
        """
        Utility class to get risk subclass used for PA_Key
        """

        if 'NT' in string:
            return 'NT'
        elif 'SPT' in string or 'TOB' in string or 'T' in string:
            return 'T'
        else:
            raise ValueError(
                "get_risk_subclass_2 : Please check the sheet name of input file. Program Terminated with value: " + string)

    def set_input_file(self, input_file):
        """
        Setter method for input_file
        """
        self.input_file = input_file

    def set_product_name(self, product_name):
        """
        Setter method for product_name
        """
        self.product_name = product_name

    def set_first_row(self, first_row):
        """
        Setter method for first_row
        """
        self.first_row = first_row

    def set_output_column_names(self, output_column_names):
        """
        Setter method for output_column_names
        """
        self.output_column_names = output_column_names


class DividendParser(BaseParser):
    """
    Dividend parser class
    """
    # Default value for row number of worksheet data including headers
    FIRST_ROW_DEFAULT = 6
    # Default value for output column names
    COLUMN_NAMES_DEFAULT = ['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Iss. Age',
                            'PA_KEY', 'CODE'] + ['Dur.' + str(i) for i in range(0, 122)]

    def __init__(self, input_file, product_name, first_row=FIRST_ROW_DEFAULT):

        self.input_file = input_file
        self.first_row = first_row
        self.product_name = product_name

    def _get_dividend_market(self, string):
        """
        Utility method to get dividend market
        """

        if 'Qual' in string:
            return 'Q'
        else:
            return 'NQ'

    def _get_dividend_type(self, string):
        """
        Utility method to get dividend type
        """

        type_list = {
            'DIV': 'Base',
            'PUA': 'PUA',
            'RPU': 'RPU',
            'LISR': 'LISR',
            'ALIR': 'ALIR'
        }

        string = string.upper()

        for type in type_list:
            if type in string:
                return type_list[type]

        raise ValueError(
            "get_dividend_type: Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_risk_class(self, string):
        """
        Utility method to get risk class
        SPT and TOB are different from base parser
        """

        class_list = {
            'UPNT': 'UPNT',
            'SPNT': 'SPNT',
            'NT': 'NT',
            'SPT': 'ST',
            'TOB': 'T'
        }

        string = string.upper()

        for type in class_list:
            if type in string:
                return class_list[type]

        raise ValueError(
            "get_risk_class: Please check the sheet name of input file. Program Terminated with value: " + string)


    def parse(self):
        """
        Parse method
        :return: Data frame
        """

        # Output data frame
        output = pd.DataFrame([])
        xl = pd.ExcelFile(self.input_file)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
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
            output = pd.concat([output, df], sort=False)

        output['Product'] = self.product_name

        # Add rate for ALIR PUA and set the value equal to PUA
        df_alirpua = output[output['Base/PUA/RPU'] == 'PUA'].copy()
        df_alirpua['Base/PUA/RPU'] = 'ALIR PUA'
        output = pd.concat([output, df_alirpua], sort=False)

        # Add rate for Qualified and set the value equal to None Qualified for ALIR, ALIR PUA, PUA, and LISR
        df_qualified = output[
            output['Base/PUA/RPU'].isin(['ALIR', 'ALIR PUA', 'PUA', 'LISR']) & (output['Gender'] == 'U')].copy()
        df_qualified['Market'] = 'Q'
        output = pd.concat([output, df_qualified], sort=False)

        # Reset index before dropping
        output = output.reset_index(drop=True)

        # Drop Qualified and Age < 17
        output.drop(output[(output['Market'] == 'Q') & (output['Age'] < 17)].index, inplace=True)

        # Drop risk class (T, ST) and Age < 15
        output.drop(output[output['Underwriting Class'].isin(['T', 'ST']) & (output['Age'] < 15)].index, inplace=True)

        # Build PA_KEY
        output['PA_KEY'] = 'CP' + output['Product'] + 'A,' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + \
                           output['Market_in_key'] + ',' + \
                           output['risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output[
                               'Band'] + ',' + output['Age'].astype(str)
        # Build CODE
        output['CODE'] = output['Product'] + ',' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output[
            'Market'] + ',' + \
                         output['Underwriting Class'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
        output['Dur.0'] = 0
        output.columns = output.columns.astype(str)
        # Rearrange the column
        output = output[
            ['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Age', 'PA_KEY', 'CODE',
             'Dur.0'] + [str(i) for i in range(1, 122)]]
        # Rename the column
        output.columns = self.COLUMN_NAMES_DEFAULT

        return output


class CurrPremPerkParser(BaseParser):
    """
    CurrPremPerK parser
    """

    # Input file worksheet names
    # LP10 HECV:  Prem Male, Prem Female, Prem Unisex (No banding)
    # LP15 20 65: Prem Male Band 2 3 4 5, Prem Female Band 2 3 4 5, Prem Unisex Band 2 3 4 5
    # LP100:      Prem Male Band 1 2 3 4 5, Prem Female Band 1 2 3 4 5, Prem Unisex Band 1 2 3 4 5
    # Sub Classes: Sub_Classified_Prem_Male_NT, Sub_Classified_Prem_Male_TOB, Sub_Classified_Prem_Female_NT,
    #              Sub_Classified_Prem_Female_TOB, Sub_Classified_Prem_Unisex_NT, Sub_Classified_Prem_Unisex_TOB


    FIRST_ROW_DEFAULT = 4
    FIRST_ROW_SUB_DEFAULT = 4
    COLUMN_NAMES_DEFAULT = ['Product', 'Gender', 'Band', 'Class', 'Table Rating', 'Issue Age', 'PA_Key', 'Code',
                            'Premium_Rate']
    RISK_CLASS_DICT = {'1': 'UPNT', '2': 'SPNT', '3': 'NT', '4': 'ST', '5': 'T'}
    RISK_CLASS_1_DICT = {'1': 'UP', '2': 'SP', '3': '', '4': 'SP', '5': ''}
    RISK_CLASS_2_DICT = {'1': 'NT', '2': 'NT', '3': 'NT', '4': 'T', '5': 'T'}
    TABLE_RATINGS = ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'J', 'L', 'P']

    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row

    def parse(self):
        """
        Parse method
        :return: Data frame
        """

        # Output data frame
        output = pd.DataFrame([])
        # Read excel file into pandas data object
        xl = pd.ExcelFile(self.input_file)
        # Find general risk classes sheets
        prem_sheet_list = [s for s in xl.sheet_names if 'Prem' == s[:4]]
        # Find sub risk classes sheets
        prem_sub_classified_sheet_list = [s for s in xl.sheet_names if 'Sub' == s[:3]]

        # 1. General risk classes, worksheet name starting with Prem #
        for sheet in prem_sheet_list:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            # Get gender and risk class information from the header <Age M1	M2	M3	M4	M5	M0>
            gender_and_risk_class = df.columns[1:6].tolist()
            # Restructure and make <M1 M2 M3 M4 M5> as dataframe value
            df = pd.melt(df, id_vars=['Age'], value_vars=gender_and_risk_class)
            # Split M1 into gender and risk class
            risk_class_info = df['variable'].str.split(pat=r'([A-Za-z]+)', expand=True).iloc[:, 1:3]
            # Split worksheet name
            info = sheet.split(' ')
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(info[1])
            # Get band from worksheet name if Band exists
            if 'Band' in sheet:
                df['Band'] = 'B' + info[-1]
            # For LP10 and HECV
            else:
                # replicate 4 times with band 2,3,4,5
                df_1 = pd.DataFrame([])
                for b in range(2, 6):
                    df['Band'] = 'B' + str(b)
                    df_1 = pd.concat([df_1,df.copy()])
                df = df_1

            df['Class'] = risk_class_info[2]
            # UP or SP
            df['risk_class_in_key_1'] = risk_class_info[2]
            # NT or T
            df['risk_class_in_key_2'] = risk_class_info[2]
            # Default table rating as -
            df['Table Rating'] = "-"
            # Concatenate to the output data frame
            output = pd.concat([output, df], sort=False)
        # Set product name
        output['Product'] = self.product_name
        # Map risk class to UP/SP, NT/T
        output = output.replace({'Class': self.RISK_CLASS_DICT, 'risk_class_in_key_1': self.RISK_CLASS_1_DICT,
                                 'risk_class_in_key_2': self.RISK_CLASS_2_DICT})
        # Generate PA_Key
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] + "," + output['Band'] + "," + output[
            'risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Age'].astype(str)
        # Generate Code
        output['Code'] = output['Product'] + ',' + output['Gender'] + ',' + output['Band'] + ',' + output[
            'Class'] + ',' + output['Table Rating'] + ',' + output['Age'].astype(str)
        # Rename columns
        output = output.rename(columns={'value': 'Premium_Rate', 'Age': 'Issue Age'})
        # Drop extra columns
        output = output.drop(columns=['risk_class_in_key_1', 'risk_class_in_key_2', 'variable'])
        # Reset index
        output = output.reset_index(drop=True)

        # 2. Sub risk classes, worksheet name starting with Sub_Classified
        output_sub = pd.DataFrame([])
        for sheet in prem_sub_classified_sheet_list:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.FIRST_ROW_SUB_DEFAULT - 1, encoding='utf8')
            # Restructure and make ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'J', 'L', 'P'] as dataframe value
            df = pd.melt(df, id_vars=['Age'], value_vars=self.TABLE_RATINGS)
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(sheet)
            # Get risk class from worksheet name
            df['Class'] = self._get_risk_class(sheet)
            # Set default empty band
            df['Band'] = ''
            # Concatenate to the output data frame
            output_sub = pd.concat([output_sub, df])
        # Set product name
        output_sub['Product'] = self.product_name
        # Rename columns
        output_sub = output_sub.rename(
            columns={'variable': 'Table Rating', 'value': 'Premium_Rate', 'Age': 'Issue Age'})
        # Generate PA_Key
        output_sub['PA_Key'] = 'CP' + output_sub['Product'] + 'A,' + output_sub['Gender'] + ',' + output_sub[
            'Class'] + ',' + output_sub['Table Rating'] + ',' + output_sub['Issue Age'].astype(str)
        # Generate Code
        output_sub['Code'] = output_sub['Product'] + ',' + output_sub['Gender'] + ',' + output_sub['Band'] + "," + \
                             output_sub[
                                 'Class'] + ',' + output_sub['Table Rating'] + ',' + output_sub['Issue Age'].astype(str)
        # Reset index
        output_sub = output_sub.reset_index(drop=True)

        # Concatenate to output
        output = pd.concat([output, output_sub], sort=False)
        # Rearrange columns
        output = output[self.COLUMN_NAMES_DEFAULT]
        # Reset index
        output = output.reset_index(drop=True)
        return output


class WaiverPerKParser(BaseParser):
    """
    WaiverPreK parser
    """

    # Input file worksheet names
    # LP10 HECV:  WP Male, WP Female, WP Unisex (No banding)
    # LP15 20 65: WP Male Band 2 3 4 5, WP Female Band 2 3 4 5, WP Unisex Band 2 3 4 5 (No band 1)
    # LP100:      WP Male Band 1 2 3 4 5, WP Female Band 1 2 3 4 5, WP Unisex Band 1 2 3 4 5


    FIRST_ROW_DEFAULT = 4
    COLUMN_NAMES_DEFAULT = ['Product', 'Gender', 'Band', 'Class', 'Issue Age', 'PA_Key', 'Code',
                            'Premium_Rate']
    RISK_CLASS_DICT = {'1': 'UPNT', '2': 'SPNT', '3': 'NT', '4': 'ST', '5': 'T'}
    RISK_CLASS_1_DICT = {'1': 'UP', '2': 'SP', '3': '', '4': 'SP', '5': ''}
    RISK_CLASS_2_DICT = {'1': 'NT', '2': 'NT', '3': 'NT', '4': 'T', '5': 'T'}

    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row


    def parse(self):

        # Output data frame
        output = pd.DataFrame([])
        # Read excel file into pandas data object
        xl = pd.ExcelFile(self.input_file)
        # Find waiver worksheets, 'WP' as prefix of worksheet name
        waiver_sheet_list = [s for s in xl.sheet_names if 'WP' == s[:2]]
        # Loop each waiver worksheet
        for sheet in waiver_sheet_list:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            # Get gender and risk class information from the header <Age M1	M2	M3	M4	M5	M0>
            gender_and_risk_class = df.columns[1:6].tolist()
            # Restructure and make <M1 M2 M3 M4 M5> as dataframe value
            df = pd.melt(df, id_vars=['Age'], value_vars=gender_and_risk_class)
            # Split M1 into gender and risk class
            risk_class_info = df['variable'].str.split(pat=r'([A-Za-z]+)', expand=True).iloc[:, 1:3]
            # Split worksheet name
            info = sheet.split(' ')
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(info[1])
            # Get band from worksheet name if Band exists
            if 'Band' in sheet:
                df['Band'] = 'B' + info[-1]
            # For LP10 and HECV
            else:
                # replicate 4 times with band 2,3,4,5
                df_1 = pd.DataFrame([])
                for b in range(2, 6):
                    df['Band'] = 'B' + str(b)
                    df_1 = pd.concat([df_1,df.copy()])
                df = df_1

            df['Class'] = risk_class_info[2]
            # UP or SP
            df['risk_class_in_key_1'] = risk_class_info[2]
            # NT or T
            df['risk_class_in_key_2'] = risk_class_info[2]
            # Concatenate to the output data frame
            output = pd.concat([output, df], sort=False)
        # Set product name
        output['Product'] = self.product_name
        # Map risk class to UP/SP, NT/T
        output = output.replace({'Class': self.RISK_CLASS_DICT, 'risk_class_in_key_1': self.RISK_CLASS_1_DICT,
                                 'risk_class_in_key_2': self.RISK_CLASS_2_DICT})
        # Generate PA_Key
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] + "," + output['Band'] + "," + output[
            'risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Age'].astype(str)

        # Generate Code
        output['Code'] = output['Product'] + ',' + output['Gender'] + ',' + output['Band'] + ',' + output[
            'Class'] + ',' + output['Age'].astype(str)
        # Rename columns
        output = output.rename(columns={'value': 'Premium_Rate', 'Age': 'Issue Age'})
        # Drop extra columns
        output = output.drop(columns=['risk_class_in_key_1', 'risk_class_in_key_2', 'variable'])
        # Rearrange columns
        output = output[self.COLUMN_NAMES_DEFAULT]
        # Reset index
        output = output.reset_index(drop=True)

        return output


class NSPParser(BaseParser):
    """
    NSP parser class
    """
    # Input file worksheet names
    # LP10 HECV:  Prem Male, Prem Female, Prem Unisex (No banding)
    # LP15 20 65: Prem Male Band 2 3 4 5, Prem Female Band 2 3 4 5, Prem Unisex Band 2 3 4 5
    # LP100:      Prem Male Band 1 2 3 4 5, Prem Female Band 1 2 3 4 5, Prem Unisex Band 1 2 3 4 5
    # Sub Classes: Sub_Classified_Prem_Male_NT, Sub_Classified_Prem_Male_TOB, Sub_Classified_Prem_Female_NT,
    #              Sub_Classified_Prem_Female_TOB, Sub_Classified_Prem_Unisex_NT, Sub_Classified_Prem_Unisex_TOB


    FIRST_ROW_DEFAULT = 4
    COLUMN_NAMES_DEFAULT = ['Product', 'Gender', 'Band', 'Class', 'Issue Age', 'PA_Key', 'Code',
                            'Premium_Rate']
    RISK_CLASS_DICT = {'1': 'UPNT', '2': 'SPNT', '3': 'NT', '4': 'ST', '5': 'T'}
    RISK_CLASS_1_DICT = {'1': 'UP', '2': 'SP', '3': '', '4': 'SP', '5': ''}
    RISK_CLASS_2_DICT = {'1': 'NT', '2': 'NT', '3': 'NT', '4': 'T', '5': 'T'}

    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row


class BOYStateReserveParser(BaseParser):
    """
    BOYStateReserve parser class
    """

    COLUMN_NAMES_DEFAULT = ['Product', 'Gender',  'Iss. Age', 'PA_Key', 'Code'] + [
                            'Dur.' + str(i) for i in range(-1, 122)]

    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row


    def parse(self):

        # Output data frame
        output = pd.DataFrame([])
        # Read excel file into pandas data object
        xl = pd.ExcelFile(self.input_file)
        # Loop each reserve worksheet
        for sheet in xl.sheet_names:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(sheet)

            # Concatenate to the output data frame
            output = pd.concat([output, df], sort=False)
        # Set product name
        output['Product'] = self.product_name
        # Add two place holder columns
        output['-1'] = 0
        output['0'] = 0
        # Generate PA_Key
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] +  ',' + output['Age'].astype(str)

        # Generate Code
        output['Code'] = output['Product'] + ',' + output['Gender'] + ',' + output['Age'].astype(str)
        # Convert column names into string for the further selection
        output.columns = output.columns.astype(str)
        # Rearrange columns
        output = output[['Product', 'Gender', 'Age', 'PA_Key', 'Code'] + [str(i) for i in range(-1, 122)]]
        # Rename column name
        output.columns = self.COLUMN_NAMES_DEFAULT
        # Reset index
        output = output.reset_index(drop=True)

        return output


class CashValuePerKParser(BaseParser):
    """
    CashValuePerK parser class
    """


    # Input file worksheet names
    # LP10:   GCV (BOY) Male, GCV (BOY) Female, GCV (BOY) Unisex
    # LP15 20 65 100: Male CV(BOY), Female CV(BOY), Unisex CV(BOY)
    # HECV:       Male UPNT CV (BOY), Male SPNT CV (BOY), Male NT CV (BOY), Male SPT CV (BOY), Male SPT CV (BOY)
    #           , Female UPNT CV (BOY), Female SPNT CV (BOY), Female NT CV (BOY), Female SPT CV (BOY), Female TOB CV (BOY)
    #           , Unisex UPNT CV (BOY), Unisex SPNT CV (BOY), Unisex NT CV (BOY), Unisex SPT CV (BOY), Unisex TOB CV (BOY)

    COLUMN_NAMES_DEFAULT = ['Product', 'Gender', 'Class', 'Iss. Age', 'PA_Key', 'Code'] + [
                            'Dur.' + str(i) for i in range(0, 122)]
    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row


    def _get_risk_class(self, string):
        """
        Utility method to get risk class
        SPT and TOB are different from base parser
        """

        class_list = {
            'UPNT': 'UPNT',
            'SPNT': 'SPNT',
            'NT': 'NT',
            'SPT': 'SPT',
            'TOB': 'TOB'
        }

        string = string.upper()

        for type in class_list:
            if type in string:
                return class_list[type]

        raise ValueError(
            "get_risk_class: Please check the sheet name of input file. Program Terminated with value: " + string)

    def parse(self):

        # Output data frame
        output = pd.DataFrame([])
        # Read excel file into pandas data object
        xl = pd.ExcelFile(self.input_file)
        # Find cash value worksheets, worksheet name contains CV (BOY) or CV(BOY)
        cash_value_sheet_list = [s for s in xl.sheet_names if 'CV (BOY)' in s or 'CV(BOY)' in s]
        # Loop each cash value worksheet
        for sheet in cash_value_sheet_list:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(sheet)
            # Get risk class for L85 only
            if self.product_name == 'L85':
                df['Class'] = self._get_risk_class(sheet)
            else:
                df['Class'] = ''
            # Concatenate to the output data frame
            output = pd.concat([output, df], sort=False)
        # Set product name
        output['Product'] = self.product_name
        # Add column 122 for L10 only
        if self.product_name == 'L10':
            output['122'] = 0
            # Set 1000 for Age 0 duration 122
            output.loc[(output.Age == 0), '122'] = 1000.00

        # Generate PA_Key
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] + ',' + output['Class'] + ',' + output[
            'Age'].astype(str)
        # Generate Code
        output['Code'] = output['Product'] + ',' + output['Gender'] + ',' + output[
            'Class'] + ',' + output['Age'].astype(str)

        # Convert column names into string for the further selection
        output.columns = output.columns.astype(str)
        # Rearrange columns
        # L10 has different header for rates from 1 to 122
        if self.product_name in ['L10', 'L12']:
            output = output[['Product', 'Gender', 'Class', 'Age', 'PA_Key', 'Code'] + [str(i) for i in range(1, 123)]]
        # Rest of product's header for rates are from 0 to 121
        else:
            output = output[['Product', 'Gender', 'Class', 'Age', 'PA_Key', 'Code'] + [str(i) for i in range(0, 122)]]
        # Rename column name
        output.columns = self.COLUMN_NAMES_DEFAULT
        # Reset index
        output = output.reset_index(drop=True)

        return output

class TAI_TRParser(BaseParser):
    """
    CashValuePerK parser class
    """

    COLUMN_NAMES_DEFAULT = ['PA_Key', 'Product', 'Gender', 'Smk Stat', 'Age', 'Code'] + [
                            'Dur.' + str(i) for i in range(1, 122)]


    def __init__(self, input_file, product_name, first_row):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row

    def _get_risk_class(self, string):
        class_list = {
            'UPNT': 'UPNT',
            'SPNT': 'SPNT',
            'NT': 'NT',
            'NS': 'NT',
            'SPT': 'SPT',
            'TOB': 'T',
            'SM': 'T'
        }

        string = string.upper()

        for type in class_list:
            if type in string:
                return class_list[type]

        raise ValueError(
            "get_risk_class: Please check the sheet name of input file. Program Terminated with value: " + string)

    def _get_risk_class_1(self, string):
        class_list = {
            'UPNT': 'N',
            'SPNT': 'N',
            'NT': 'N',
            'NS': 'N',
            'SPT': 'S',
            'TOB': 'S',
            'SM': 'S'
        }

        string = string.upper()

        for type in class_list:
            if type in string:
                return class_list[type]

        raise ValueError(
            "get_risk_class: Please check the sheet name of input file. Program Terminated with value: " + string)

    def parse(self):

        # Output data frame
        output = pd.DataFrame([])
        # Read excel file into pandas data object
        xl = pd.ExcelFile(self.input_file)
        # Loop each TAI_TR worksheet
        for sheet in xl.sheet_names:
            # Convert each worksheet into dataframe
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            # Get gender from worksheet name
            df['Gender'] = self._get_gender(sheet)
            # Get risk class from worksheet name
            df['Smk Stat'] = self._get_risk_class(sheet)
            # Get risk class used for Code
            df['Code_Class'] = self._get_risk_class_1(sheet)
            # Concatenate to the output data frame
            output = pd.concat([output, df], sort=False)
        # Set product name
        output['Product'] = self.product_name
        # Generate PA_Key
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] + ',' + output['Smk Stat'] + ','+ output[
            'Age'].astype(str)
        # Generate Code
        output["Code"] = output["Product"] + "," + output["Gender"] + "," + output["Code_Class"] + "," + output[
        "Age"].astype(str)
        output.columns = output.columns.astype(str)

        # Convert column names into string for the further selection
        output.columns = output.columns.astype(str)
        # Rearrange columns
        output = output[['PA_Key', 'Product', 'Gender', 'Smk Stat', 'Age', 'Code'] + [str(i) for i in range(1, 122)]]
        # Rename column name
        output.columns = self.COLUMN_NAMES_DEFAULT
        # Reset index
        output = output.reset_index(drop=True)

        return output


def parser_factory(parserType):
    """
    Factory function to return parser class based on input string
    :param parserType: Parser type string
    :return: Parser class
    """

    parsers = {
        'Dividend': DividendParser,
        'CurrPremPerK': CurrPremPerkParser,
        'WaiverPerK': WaiverPerKParser,
        'NSP': NSPParser,
        'BOYStateReserve': BOYStateReserveParser,
        'CashValuePerK': CashValuePerKParser,
        'TAI_TR': TAI_TRParser
    }

    return parsers[parserType]


def validation(srcFile, destnFile):
    """
    Compare two csv files
    :param srcFile: source csv file path
    :param destnFile: destination csv file path
    :return: Ture if all matches, False otherwise
    """

    df_src = pd.read_csv(srcFile).fillna(0)
    df_destn = pd.read_csv(destnFile).fillna(0)
    df_diff = df_src - df_destn
    return (df_diff == 0).any().any()


def get_product(string):
    """
    Get product name based on input string
    :param string:
    :return:
    """

    product_list = {
        'LP10': 'L10',
        'LP12': 'L12',
        'LP15': 'L15',
        'LP20': 'L20',
        'LP65': 'L65',
        'HECV': 'L85',
        'L100': 'L100'
    }

    for product in product_list:
        if product in string:
            return product_list[product]

    raise ValueError("get_product : Please check the name of input file. Program Terminated with value: " + string)


def main():
    """
    Main application function
    :return:
    """

    # Load configuration file
    config = configparser.ConfigParser()
    config.read('config.txt')
    # Get parser type through command line
    if len(sys.argv) >= 2:
        parser_type = sys.argv[1]

    # parser_type = 'Dividend'
    # parser_type = 'CurrPremPerK'
    # parser_type = 'WaiverPerK'
    # parser_type = 'CashValuePerK'
    # parser_type = 'BOYStateReserve'
    parser_type = 'TAI_TR'


    # Load input and output config
    io_dic = config['IO']
    if parser_type == 'Dividend':
        input_dir = io_dic['Dividend.input_dir']
    elif parser_type == 'BOYStateReserve':
        input_dir = io_dic['Reserve.input_dir']
    elif parser_type == 'TAI_TR':
        input_dir = io_dic['TAI_TR.input_dir']
    else:
        input_dir = io_dic['Rate.input_dir']

    output_file = io_dic['Output_file']

    parser_config = config[parser_type]

    df_output = pd.DataFrame([])
    input_file_list = listdir(input_dir)

    for eachFile in input_file_list:
        input_file = input_dir + '\\' + eachFile
        product_name = get_product(eachFile)
        first_row = parser_config[f"{product_name}.data_first_row"]
        # Initialize a parser
        parser = parser_factory(parser_type)(input_file, product_name, int(first_row))
        df_output = pd.concat([df_output, parser.parse()], sort=False)

    # Fill NaN value with 0
    df_output = df_output.fillna(0)
    # Reset Index
    df_output = df_output.reset_index(drop=True)

    # Write data frame into excel file with proper sheet name
    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name=parser_config['Output_sheet_name'], index=False, startcol=1)

    # writer = pd.ExcelWriter(output_file, mode='a', engine='openpyxl')
    # df_output.to_excel(writer, sheet_name=parser_config['Output_sheet_name'], index=False, startcol=1)
    # writer.save()


if __name__ == '__main__':
    main()
