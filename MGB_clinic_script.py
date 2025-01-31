import xlwings as xw # must install pip, and install xlwings using pip
from datetime import datetime, timedelta

MONTH_NUM = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}

MGB_EARS_file_path = "Internal Medicine EARs AY25 MGB.xlsm"
MGB_Clinic_file_path = "MGB IM Clinic SCHEDULE AY24.xlsx"
wb = xw.Book(MGB_EARS_file_path)
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
    wb.save(MGB_EARS_file_path)

# ***************************************************************
# #algorithm for extracting info from MGB IM Clinic Schedule
# ***************************************************************

mgb_clinic_book = xw.Book(MGB_Clinic_file_path)
mgb_clinic_sheet = mgb_clinic_book.sheets["VA Clinic Report"]
row_count = 2
while mgb_clinic_sheet.range("A" + str(row_count)).value != None:
    row_count+=2
row_count -=1

# note: switching to (row_num, col_num) for cell coordinates 
col_count = 2
while mgb_clinic_sheet.range((1, col_count)).value != None or mgb_clinic_sheet.range((1, col_count+1)).value != None:
    col_count+=1

for row in range(2, row_count+1):
    shift_type = ''
    if row%2 == 0: # even rows = first shift for an individual
        name = mgb_clinic_sheet.range((row, 1)).value
        shift_type = 'AM'
    else: # odd rows = second shift for an individual 
        name = mgb_clinic_sheet.range((row-1, 1)).value
        shift_type = 'PM'
    if "(" in name:
        name = name.split(' (')[0] # removes parenthesis to prevent indexing errors
    for col in range(2, col_count+1):
        if mgb_clinic_sheet.range((row,col)).value != None: # if the cell is not blank, we want to fill in the corresponding shift on spreadsheet
            # get the date
            date = mgb_clinic_sheet.range((1, col)).value
            year = date.year
            month = date.month
            day = date.day
            EARS_sheet = get_sheet(MGB_EARS_file_path, year, month, day)
            if EARS_sheet is None:
                print('the date:', date, 'does not exist on this sheet.', 'name:', name)
            else:
                (cell_cor, sheet) = get_cell(EARS_sheet, day, shift_type, name)
                if cell_cor is None: # if name not found on spreadsheet, get_cell
                    # returns (None, None)
                    print(name, 'not found on spreadsheet :(')
                    continue
                set_cell_present(cell_cor, sheet)
