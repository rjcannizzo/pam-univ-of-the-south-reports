"""
Explore the input files to determine how to create the app(s)
2021-11-07
"""
import pandas as pd
import re
import csv

dq_file = r"D:\Python\projects\pam\2021\Excel_cleanup\project\examples\DQ Prop chart 11.5.21 - Generic Example.xlsx"
u_south_1 = r"D:\Python\projects\pam\2021\Excel_cleanup\project\examples\Univ of South Group July 2021.xlsx"
u_south_2 = r"D:\Python\projects\pam\2021\Excel_cleanup\project\examples\Univ of South Individual July 2021.xlsx"

# for sheet_name, data in file:data
file_data = {'Group Combined': None, 'McClurg': None, 'Pub': None, 'Stirlings': None, 'Cup Gown': None,	'St Andrews': None}

def get_header():
    """
    The header should have 16 columns
    """
    return ['Category','Item #','Pack','Size','Brand','Description','MPC Code','CW','Cs Qty','Cs Total $','Cs Avg $','Split Qty','Split Total $','Split Avg $','Weight','Total Sales $']
    

def column_is_integer(text):
    try:
        int(text)
    except ValueError:
        return False
    return True


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


def convert_to_csv(input_file, output_file, skiprows=7):
    """
    Convert an Excel file to a csv.
    :param skiprows: number of rows to skip when reading the Excel file.
    :param input_file: full path to the Excel file to convert.
    :param output_file: full path of the csv file to write.
    :return: path to csv file as a string?
    """
    df = pd.read_excel(input_file, skiprows=skiprows, usecols="A:N,P")
    df.to_csv(output_file, header=True, index=False, encoding='utf-8')
    return output_file


def parse_totals_file(filepath):
    """
    Returns a list of row data from the first of two files needed. This data is for tab 1 of the report.
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
                    row.insert(0, get_category(row[1]))                    
                    data.append(row)
    file_data['Group Combined'] = data

  
def write_data_to_csv(filepath, data):
    header = get_header()
    with open(filepath, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(data)


def row_is_location(text):
    # sites = {"ST. ANDREWS", "STIRLING'S COFFEE HOUSE", "CUP & GOWN", "SOUTH MCCLURG DINING", "PUB"}
    sites = {"ST. ANDREWS": "St Andrews", "STIRLING'S COFFEE HOUSE": "Stirlings", "CUP & GOWN": "Cup Gown", "SOUTH MCCLURG DINING": "McClurg", "PUB": "Pub" }
    for site in sites.keys():
        mo = re.search(site, text)
        if mo:
            return sites[site]
    return False


def parse_sites_file(filepath):
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
                    row.insert(0, get_category(row[1]))                    
                    data.append(row)

        file_data[location] = data
            

def write_group_data_to_csv():
    """This data is for tab 1; this was my initial test"""
    filepath = convert_to_csv(u_south_1, 'u_south_1.csv', skiprows=7)
    data = parse_file(filepath)
    file_data['Group Combined'] = data
    write_data_to_csv('u_south_group_output.csv', data)

def run():
    totals_file = convert_to_csv(u_south_1, 'u_south_1.csv', skiprows=7)
    parse_totals_file(totals_file)
    sites_file = convert_to_csv(u_south_2, 'u_south_2.csv', skiprows=7)
    parse_sites_file(sites_file)      
          
    with pd.ExcelWriter("group_text.xlsx") as writer: 
        for site, data in file_data.items():
            if not data:
                continue             
            df = pd.DataFrame(data, columns=get_header())
            df.to_excel(writer, sheet_name=site, index=False)

def main():
    run()


if __name__ == '__main__':
    main()
