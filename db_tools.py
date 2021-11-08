"""
Code to build the 'category' database. It's used to add a 'category' to the reports for University of the south.
2021-11-08
"""

import logging
from sqlite3.dbapi2 import Cursor
import numpy as np
import pandas as pd
from pathlib import Path
import sqlite3
from parsers import parse_int
from my_logger import logger


def iterate_folder(folder): 
    """Yield Excel files from a folder."""   
    for file in Path(folder).glob('*.xlsx'):        
        yield file


def get_dataframe_from_excel(filepath, cols='A:B'):
    """Return a dataframe from the 1st tab of an Excel file. Columns 0 and 1 converted to Integers. Rows with nan are dropped."""
    convert_to_int = lambda c: int(c)
    df = pd.read_excel(filepath, usecols=cols, converters={0: convert_to_int, 1: convert_to_int})
    df.columns = ['category', 'item']
    df.dropna(inplace=True)
    logger.info(filepath)
    return df

def insert_from_dataframe(df):    
    """Insert item and the associated category from a dataframe"""  
    conn = sqlite3.connect(r'D:\Python\projects\pam\2021\univ_south_report_tool\categories.db')
    cursor = conn.cursor()
    query = "INSERT OR IGNORE INTO category (item, category) VALUES(?,?);" 
    for row in df.itertuples(index=False, name='row'):
        category, item = row
        cursor.execute(query, (item, category))
    conn.commit()
    conn.close()
    

def main():
    db_source_folder = r'D:\Python\projects\pam\2021\univ_south_report_tool\category_database_source'
    for file in iterate_folder(db_source_folder):
        df = get_dataframe_from_excel(file)
        insert_from_dataframe(df)
    

if __name__ == '__main__':
    main()