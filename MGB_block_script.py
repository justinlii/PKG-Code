# import pandas as pd
import xlwings as xw # must install pip, and install xlwings using pip
from datetime import datetime, timedelta

MONTH_NUM = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}

MGB_EARS_file_path = "Internal Medicine EARs AY25 MGB.xlsm"
MGB_Block_file_path = "MGB IM Block SCHEDULE AY24.xlsx"
MGB_EARS_wb = xw.Book(MGB_EARS_file_path)
IRRELEVANT_SHEETS = ["<Sheet [Internal Medicine EARs AY25 MGB.xlsm]HPT_List>", 
                     "<Sheet [Internal Medicine EARs AY25 MGB.xlsm]D>",
                     "<Sheet [Internal Medicine EARs AY25 MGB.xlsm]CODES>",
                     "<Sheet [Internal Medicine EARs AY25 MGB.xlsm]FINAL RECONCILIATION>",
                     "<Sheet [Internal Medicine EARs AY25 MGB.xlsm]EAR_OVERVIEW>"] 


# useful helper functions 
def get_sheet(path, year, month, day):
    """Given file path to EARS file, year, month, and day, find the correct EARS spreadsheet within the file"""
    wb = xw.Book(path)
    relevant_sheets = []
    for sheet in wb.sheets:
        if str(sheet) not in IRRELEVANT_SHEETS:
            relevant_sheets.append(sheet)
    given_date = datetime(year, month, day)
    for sheet in relevant_sheets:
        # check year associated with sheet
        sheet_year = int(sheet.range("C9").value)
        if sheet_year == year: # Find sheet with correct year
            sheet_month = MONTH_NUM[sheet.range("C8").value]
            sheet_start_date = datetime(sheet_year, sheet_month, 1)
            sheet_end_date = sheet.range("G4").value
            if sheet_start_date <= given_date and given_date <= sheet_end_date: # Find sheet with correct date range ex: 7/1 - 7/31
                return sheet
    # if no sheet found, return none
    return None

def make_all_names_dict(path):
    """Given a file path construct and return a dictionary that holds the name 
    to row mappings for all the spreadsheets with the file"""
    wb = xw.Book(path)
    sheets_name_mapping = {}
    for sheet in wb.sheets:
        if str(sheet) not in IRRELEVANT_SHEETS:
            all_names = sheet.range("B13:B361").value[::2] # list slicing is needed because there are two rows for each name
            names_dict = {}
            counter = 13
            for name in all_names:
                names_dict[name] = "B" + str(counter)
                counter+=2
            sheets_name_mapping[str(sheet)] = names_dict
    return sheets_name_mapping

all_sheets_name_mappings = make_all_names_dict(MGB_EARS_file_path)

def get_cell(sheet, day, shift, name):
    """Given the EARS sheet, day, first_shift status, and name return a tuple of (corresponding coordinate, sheet)"""
    # check that the name is actually on the spreadsheet

    # if the name passed into this function has a title ex: "Vergara Greeno, Rebeca (DGM)"
    name = name.split(' (')[0]
    if name not in all_sheets_name_mappings[str(sheet)]:
        return (None, None) # name not found on spread sheet
    str_sheet = str(sheet)
    str_cor = ""
    if 73 + day - 1 > 90:
        str_cor += "A"
        str_cor += chr(73 + day - 1 - 90 + 64)
    else:
        str_cor += chr(73 + day - 1)
    
    if shift == 'AM':
        row_num = all_sheets_name_mappings[str_sheet][name][1:] # we only want the row number, ignore the first letter
        str_cor += row_num
    elif shift == 'PM':
        row_num = int(all_sheets_name_mappings[str_sheet][name][1:]) + 1
        str_cor += str(row_num)
    return (str_cor, sheet)


def set_cell_present(cor, sheet):
    """Given a sheet and a coordinate, set the value within that sheet to present"""
    sheet.range(cor).value = "P"
    print("set", cor, "in file:", sheet)
    MGB_EARS_wb.save(MGB_EARS_file_path)

# ***************************************************************
# algorithm for extracting data from MGB IM Block Schedule Report
# ***************************************************************
mgb_block_book = xw.Book(MGB_Block_file_path)
mgb_block_sheet = mgb_block_book.sheets[0] # Access the first sheet in the list of sheets (there is only one anyway)
# remember to increment curr_year by 1

def fill_mgb_block_sheet(block_sheet, start_row, start_col):
    """Given the block_sheet and the starting row, col, populate the corresponding EARS sheet"""
    curr_date_cell = (start_row, start_col)
    curr_year = int("20" + block_sheet.name[2:4]) # Get the year from name of the sheet of interest ("20" + "23")
    while mgb_block_sheet.range(curr_date_cell).value != None:
        #ex:  "9/20 - 10/3"
        date_range_string = mgb_block_sheet.range(curr_date_cell).value
        start_date_str, end_date_str = date_range_string.split(' - ')

        start_date = datetime.strptime(start_date_str, "%m/%d").replace(year=curr_year)
        end_date = datetime.strptime(end_date_str, "%m/%d").replace(year=curr_year)

        if end_date < start_date: # "date-range stretches across a new year"
            curr_year += 1
            end_date = end_date.replace(year=curr_year)
        
        # Generate list of all dates in the range (inclusive)
        current_date = start_date
        dates_between = [] # list stores date objects
        while current_date <= end_date:
            dates_between.append(current_date)
            current_date += timedelta(days=1)  # Increment by one day

        # getting the columns of names associated underneath a date range
        curr_name_cell = (curr_date_cell[0] + 1, curr_date_cell[1])

        while mgb_block_sheet.range(curr_name_cell).value != None:  # iterate downwards while there are still names in that column
            name = mgb_block_sheet.range(curr_name_cell).value
            if name == 'Holiday Coverage':
                curr_name_cell = (curr_name_cell[0]+1, curr_name_cell[1]) #increment row by 1
                continue
            for date in dates_between:
                EARS_sheet = get_sheet(MGB_EARS_file_path, date.year, date.month, date.day)
                # if no EARS_sheet found, means there aren't any spreadsheets with the listed date
                if EARS_sheet == None:
                    print('No EARS sheet for date:', date, '(', name, ')')
                    continue # move onto the next date

                cell_1 = get_cell(EARS_sheet, date.day, 'AM', name)
                cell_2 = get_cell(EARS_sheet, date.day, 'PM', name)

                if cell_1[0] == None:
                    print(name, 'not found on spreadsheet :(')
                else:
                    set_cell_present(*cell_1)
                if cell_2[0] == None:
                    print(name, 'not found on spreadsheet :(')
                else:
                    set_cell_present(*cell_2)
            curr_name_cell = (curr_name_cell[0]+1, curr_name_cell[1]) # increment row by 1
            print('current name cell coordinates:', curr_name_cell)
        curr_date_cell = (curr_date_cell[0], curr_date_cell[1] + 1) # increment date cell column by one
        

fill_mgb_block_sheet(mgb_block_sheet, 3, 2)
fill_mgb_block_sheet(mgb_block_sheet, 12, 3 )
