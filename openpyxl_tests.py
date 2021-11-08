"""
I'd like to find the first row of data in Excel files.
Also, how to add a table t the data on a tab.
"""
import pandas as pd
import openpyxl



def find_first_header_row(filepath):
    """
    I may need this to determine the number of rows to skip when running convert_to_csv().
    Series.str.find('Item #')
    :param filepath: full path to the Excel file
    :return: 
    """
    df = pd.read_excel(filepath, usecols=[0], squeeze=True)   
    # returns 0.0 if found - I think this is the starting point of the string 
    series = df.str.find('Item #') 
    