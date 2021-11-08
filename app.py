"""

2021-11-08
"""
import csv
import logging
import pandas as pd
import re
import sqlite3
from parsers import parse_float, parse_int, parse_string

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
s_handler = logging.StreamHandler()
f_handler = logging.FileHandler('logs/app.log')
f_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
s_handler.setFormatter(formatter)
f_handler.setFormatter(formatter)
logger.addHandler(s_handler)
logger.addHandler(f_handler)

u_south_1 = r"D:\Python\projects\pam\2021\univ_south_report_tool\examples\Univ of South Group July 2021.xlsx"
u_south_2 = r"D:\Python\projects\pam\2021\univ_south_report_tool\examples\Univ of South Individual July 2021.xlsx"

file_data = {'Group Combined': None, 'McClurg': None, 'Pub': None, 'Stirlings': None, 'Cup Gown': None,	'St Andrews': None}

def get_header():
    """
    Return a list of column headers. The header should have 16 columns.
    """
    return ['Category','Item #','Pack','Size','Brand','Description','MPC Code','CW','Cs Qty','Cs Total $','Cs Avg $','Split Qty','Split Total $','Split Avg $','Weight','Total Sales $']
    

def column_is_integer(text):
    try:
        int(text)
    except ValueError:
        return False
    return True


def convert_to_csv(input_file, output_file, skiprows=7):
    """
    Convert an Excel file to a csv.
    :param skiprows: number of rows to skip when reading the Excel file.
    :param input_file: full path to the Excel file to convert.
    :param output_file: full path of the csv file to write.
    :return: path to csv file as a string?
    """    
    converters={0: parse_int, 1: parse_int}
    df = pd.read_excel(input_file, skiprows=skiprows, usecols="A:N,P")
    df.to_csv(output_file, index=False, encoding='utf-8')
    return output_file


def find_first_header_row(filepath):
    """
    I may need this to determine the number of rows to skip when running convert_to_csv().
    Series.str.find('Item #')
    :param filepath: full path to the Excel file
    :return: 
    """
    df = pd.read_excel(filepath, usecols=[0], squeeze=True)    
    a = df.str.find('Item #') # returns 0.0 if found


def get_category(item_number):
    return 'NA'


def parse_row(row):
	"""
	Return a list of data cleaned by parsers (e.g. parse_int)		
	"""
	column_parsers = [parse_string, parse_int, parse_string, parse_string, parse_string, parse_string, parse_string,
        parse_int, parse_float, parse_float, parse_int, parse_float, parse_float, parse_float, parse_float, parse_float]
	return [func(field) for func, field in zip(column_parsers, row)]


def parse_file_2(filepath):
    location = None
    data = []
    with open(filepath, encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if not column_is_integer(row[0]):
                site = row_is_location(row[0])
                if site and data:
                    file_data[location] = data
                    data = []
                    location = site
                if site and not data:
                    location = site
                continue
            else:                
                case_quantity = float(row[7])
                split_quantity = float(row[10])
                if any([case_quantity > 0, split_quantity > 0]):
                    row = [item for item in row if not item.startswith('Unnamed')]
                    row = parse_row(row)  
                    row.insert(0, get_category(row[1]))                                      
                    data.append(row)

        file_data[location] = data


def parse_file_1(filepath):
    """
    Returns a list of row data from the first of two files needed. This is consolidated data for tab 1 of the report.
    """
    data = []
    with open(filepath, encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if row[0].startswith('Count'):
                break
            else:                
                case_quantity = float(row[7])
                split_quantity = float(row[10])
                if any([case_quantity > 0, split_quantity > 0]):
                    row = [item for item in row if not item.startswith('Unnamed')]                     
                    row = parse_row(row)     
                    row.insert(0, get_category(row[1]))              
                    data.append(row)
    file_data['Group Combined'] = data


def row_is_location(text):
    """Return True if a row is a location header row else False."""
    sites = {"ST. ANDREWS": "St Andrews", "STIRLING'S COFFEE HOUSE": "Stirlings", "CUP & GOWN": "Cup Gown", "SOUTH MCCLURG DINING": "McClurg", "PUB": "Pub" }
    for site in sites.keys():
        mo = re.search(site, text)
        if mo:
            return sites[site]
    return False


def run_report():
    """
    Manager function to run all necessary functions to write the report(s).
    """    
    totals_file = convert_to_csv(u_south_1, 'u_south_1.csv', skiprows=7)
    parse_file_1(totals_file)
    sites_file = convert_to_csv(u_south_2, 'u_south_2.csv', skiprows=7)
    parse_file_2(sites_file)      
          
    with pd.ExcelWriter("group_text.xlsx") as writer: 
        for site, data in file_data.items():
            if not data:
                continue             
            df = pd.DataFrame(data, columns=get_header())           
            df.to_excel(writer, sheet_name=site, index=False)


def main():
    run_report()


if __name__ == '__main__':
    main()