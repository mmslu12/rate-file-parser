import pandas as pd
import sys
from os import listdir
import configparser
import parser_factory

class BaseParser():

    def __init__(self, input_file, product_name, first_row):
        self.input_file = input_file
        self.first_row = first_row
        self.product_name = product_name


    def _get_gender(self,string):

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


    def _get_risk_subclass_1(self, string):

        if 'SP' in string:
            return 'SP'
        elif 'UP' in string:
            return 'UP'
        else:
            return ''


    def _get_risk_subclass_2(self, string):

        if 'NT' in string :
            return 'NT'
        elif 'SPT' in string or 'TOB' in string or 'T' in string:
            return 'T'
        else:
            raise ValueError("get_risk_subclass_2 : Please check the sheet name of input file. Program Terminated with value: " + string)


    def set_input_file(self, input_file):
        self.input_file = input_file

    def set_product_name(self, product_name):
        self.product_name = product_name

    def set_first_row(self, first_row):
        self.first_row = first_row

    def set_output_column_names(self, output_column_names):
        self.output_column_names = output_column_names

class DividendParser(BaseParser):

    FIRST_ROW_DEFAULT = 6
    SHEET_NAME_DEFAULT = '^Dividends_v7_2021_Final'
    COLUMN_NAMES_DEFAULT = ['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Iss. Age', 'PA_KEY', 'CODE'] +\
                   [ 'Dur.' + str(i) for i in range(0,122)]

    def __init__(self, input_file, product_name, first_row = FIRST_ROW_DEFAULT):

        self.input_file = input_file
        self.first_row = first_row
        self.product_name = product_name


    def _get_risk_class(self, string):

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

    def _get_dividend_market(self, string):

        if 'Qual' in string:
            return 'Q'
        else:
            return 'NQ'


    def _get_dividend_type(self, string):

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
            output = pd.concat([output, df], sort=False)

        output['Product'] = self.product_name

        # Add rate for ALIR PUA and set the value equal to PUA
        df_alirpua = output[output['Base/PUA/RPU'] == 'PUA'].copy()
        df_alirpua['Base/PUA/RPU'] = 'ALIR PUA'
        output = pd.concat([output, df_alirpua])

        # Add rate for Qualified and set the value equal to None Qualified for ALIR, ALIR PUA, PUA, and LISA
        df_qualified = output[output['Base/PUA/RPU'].isin(['ALIR', 'ALIR PUA', 'PUA', 'LISR']) & (output['Gender'] == 'U')].copy()
        df_qualified['Market Type'] = 'Q'
        output = pd.concat([output, df_qualified], sort=False)

        # Reset index before dropping
        output = output.reset_index(drop=True)

        # Drop Qualified and Age < 17
        output.drop(output[(output['Market Type'] == 'Q') & (output['Age'] < 17)].index, inplace=True)

        # Drop risk class (T, ST) and Age < 15
        output.drop(output[output['Underwriting Class'].isin( ['T', 'ST']) & (output['Age'] < 15)].index, inplace=True)

        # Build PA_KEY
        output['PA_KEY'] = 'CP' + output['Product'] + 'A,' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market_in_key'] + ',' + \
                            output['risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
        # Build CODE
        output['CODE'] = output['Product'] + ',' + output['Base/PUA/RPU'] + ',' + output['Gender'] + ',' + output['Market'] + ',' + \
                            output['Underwriting Class'] + ',' + output['Band'] + ',' + output['Age'].astype(str)
        output['Dur.0'] = 0
        output.columns = output.columns.astype(str)
        # Rearrange the column
        output = output[['Product', 'Base/PUA/RPU', 'Gender', 'Market', 'Underwriting Class', 'Band', 'Age', 'PA_KEY', 'CODE', 'Dur.0'] + [str(i) for i in range(1,122)]]
        # Rename the column
        output.columns = self.COLUMN_NAMES_DEFAULT

        return output


class CurrPremPerkParser(BaseParser):

    # Input file sheet names
    # LP10 HECV:  Prem Male, Prem Female, Prem Unisex
    # LP15 20 65: Prem Male Band 2 3 4 5, Prem Female Band 2 3 4 5, Prem Unisex Band 2 3 4 5
    # LP100:      Prem Male Band 1 2 3 4 5, Prem Female Band 1 2 3 4 5, Prem Unisex Band 1 2 3 4 5
    # Sub Classes: Sub_Classified_Prem_Male_NT, Sub_Classified_Prem_Male_TOB, Sub_Classified_Prem_Female_NT,
    #              Sub_Classified_Prem_Female_TOB, Sub_Classified_Prem_Unisex_NT, Sub_Classified_Prem_Unisex_TOB

    FIRST_ROW_DEFAULT = 4
    FIRST_ROW_SUB_DEFAULT = 4
    SHEET_NAME_DEFAULT = '^CurrPremPerK_v7_2021_Final'
    COLUMN_NAMES_DEFAULT = ['Product', 'Gender', 'Band', 'Class', 'Table Rating', 'Issue Age', 'PA_Key', 'Code', 'Premium_Rate']
    RISK_CLASS_DICT = {'1': 'UPNT', '2': 'SPNT', '3': 'NT', '4': 'ST', '5': 'T'}
    RISK_CLASS_1_DICT = {'1': 'UP', '2': 'SP', '3': '', '4': 'SP', '5': ''}
    RISK_CLASS_2_DICT = {'1': 'NT', '2': 'NT', '3': 'NT', '4': 'T', '5': 'T'}
    TABLE_RATINGS = ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'J', 'L', 'P']

    def __init__(self, input_file, product_name, first_row=FIRST_ROW_DEFAULT):

        self.input_file = input_file
        self.product_name = product_name
        self.first_row = first_row

    def parse(self):

        output = pd.DataFrame([])
        xl = pd.ExcelFile(self.input_file)
        prem_sheet_list = [s for s in xl.sheet_names if 'Prem' == s[:4]]
        prem_sub_classified_sheet_list = [s for s in xl.sheet_names if 'Sub' == s[:3]]

        # General risk classes
        for sheet in prem_sheet_list:
            df = xl.parse(sheet_name=sheet, skiprows=self.first_row - 1, encoding='utf8')
            gender_and_risk_class = df.columns[1:6].tolist()
            df = pd.melt(df, id_vars=['Age'], value_vars=gender_and_risk_class)
            risk_class_info = df['variable'].str.split(pat=r'([A-Za-z]+)', expand=True).iloc[:, 1:3]
            info = sheet.split(' ')
            df['Gender'] = self._get_gender(info[1])
            df['Band'] = 'B' + info[-1] if 'Band' in sheet else ''
            df['Class'] = risk_class_info[2]
            df['risk_class_in_key_1'] = risk_class_info[2]
            df['risk_class_in_key_2'] = risk_class_info[2]
            df['Table Rating'] = "-"
            output = pd.concat([output, df], sort=False)
        output['Product'] = self.product_name
        output = output.replace({'Class': self.RISK_CLASS_DICT, 'risk_class_in_key_1': self.RISK_CLASS_1_DICT, 'risk_class_in_key_2': self.RISK_CLASS_2_DICT})
        output['PA_Key'] = 'CP' + output['Product'] + 'A,' + output['Gender'] + "," + output['Band'] + "," + output[
            'risk_class_in_key_1'] + ',' + output['risk_class_in_key_2'] + ',' + output['Age'].astype(str)
        output['Code'] = output['Product'] + ',' + output['Gender'] + ',' + output['Band'] + ',' + output[
            'Class'] + ',' + output['Table Rating'] + ',' + output['Age'].astype(str)
        output = output.rename(columns={'value': 'Premium_Rate', 'Age': 'Issue Age'})
        output = output.drop(columns=['risk_class_in_key_1', 'risk_class_in_key_2', 'variable'])

        # Sub risk classes
        output_sub = pd.DataFrame([])
        for sheet in prem_sub_classified_sheet_list:
            df = xl.parse(sheet_name=sheet, skiprows=self.FIRST_ROW_SUB_DEFAULT - 1, encoding='utf8')
            df = pd.melt(df, id_vars=['Age'], value_vars=self.TABLE_RATINGS)
            df['Gender'] = self._get_gender(sheet)
            df['Class'] = self._get_risk_class(sheet)
            df['Band'] = ''
            output_sub = pd.concat([output_sub, df])
        output_sub['Product'] = self.product_name
        output_sub = output_sub.rename(
            columns={'variable': 'Table Rating', 'value': 'Premium_Rate', 'Age': 'Issue Age'})
        output_sub['PA_Key'] = 'CP' + output_sub['Product'] + 'A,' + output_sub['Gender'] + ',' + output_sub['Class'] + ',' + output_sub[
            'Table Rating'] + ',' + output_sub['Issue Age'].astype(str)
        output_sub['Code'] = output_sub['Product'] + ',' + output_sub['Gender'] + ',' + output_sub['Band'] + "," + output_sub[
            'Class'] + ',' + output_sub['Table Rating'] + ',' + output_sub['Issue Age'].astype(str)

        output_sub = output_sub.reset_index(drop=True)
        output = pd.concat([output, output_sub], sort=False)
        output = output[self.COLUMN_NAMES_DEFAULT]
        output = output.reset_index(drop=True)
        return output


class WaiverPerkParser(BaseParser):
    def __init__(self):
        pass


class NSPParser(BaseParser):
    def __init__(self):
        pass


class BOYStateReserveParser(BaseParser):
    def __init__(self):
        pass


class CashValuePerKParser(BaseParser):
    def __init__(self):
        pass

def parser_factory(parserType):

    parsers = {
        'Dividend': DividendParser,
        'CurrPremPerK': CurrPremPerkParser,
        'Waiver': WaiverPerkParser,
        'NSP': NSPParser,
        'BOYStateReserve': BOYStateReserveParser,
        'CashValue': CashValuePerKParser
    }

    return parsers[parserType]


def validation(srcFile, destnFile):
    '''
    Compare two csv files
    :param srcFile: source csv file path
    :param destnFile: destination csv file path
    :return: Ture if all matches, False otherwise
    '''

    df_src = pd.read_csv(srcFile).fillna(0)
    df_destn = pd.read_csv(destnFile).fillna(0)
    df_diff = df_src - df_destn
    return(df_diff == 0).any().any()


def get_product(string):

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

    # Load configuration file
    config = configparser.ConfigParser()
    config.read('config.txt')
    # Get parser type through command line
    if len(sys.argv) >= 2:
        parser_type = sys.argv[1]

    # parser_type = 'Dividend'
    parser_type = 'CurrPremPerK'

    # Load input and output config
    io_dic = config['IO']
    input_dir = io_dic[f"{'Dividend' if parser_type == 'Dividend' else 'Rate'}.input_dir"]
    output_file = io_dic['Output_file']

    parser_config = config[parser_type]

    ##############################
    ####    ^Dividends    ####
    ##############################

    df_output = pd.DataFrame([])
    input_file_list = listdir(input_dir)

    for eachFile in input_file_list:

        input_file = input_dir+'\\'+eachFile
        product_name = get_product(eachFile)
        first_row = parser_config[f"{product_name}.data_first_row"]
        # Initialize a parser
        parser = parser_factory(parser_type)(input_file, product_name, int(first_row))
        df_output = pd.concat([df_output, parser.parse()], sort=False)

    # Fill NaN value with 0
    df_output = df_output.fillna(0)
    # Reset Index
    df_output = df_output.reset_index(drop=True)

    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name=parser_config['Output_sheet_name'], index=False, startcol=1)

    ##############################################################
    ####    ^CurrPremPerK, ^CashValuePerK, ^WaiverPerK, ^NSP  ####
    ##############################################################
    #
    # input_dir_rate_file = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\Rate Files'
    # output_file = 'C:\\Users\\mm13825\\OneDrive - MassMutual\\MyDocuments\\Life\\Mini Project\\Rates Files Conversion\\CurrPremPerK.xlsx'
    #
    # df_output = pd.DataFrame([])
    # input_file_list = listdir(input_dir_rate_file)
    #
    # for eachFile in input_file_list:
    #     parser = CurrPremPerkParser(input_file=input_dir_rate_file+'\\'+eachFile)
    #     df_output = pd.concat([df_output, parser.parse()])
    #
    # df_output = df_output.reset_index(drop=True)
    #
    # with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
    #     df_output.to_excel(writer, sheet_name=CurrPremPerkParser.SHEET_NAME_DEFAULT)

if __name__ == '__main__':
    main()