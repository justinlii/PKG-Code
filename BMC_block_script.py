import xlwings as xw # must install pip before installing xlwings
from datetime import datetime, timedelta

# # BMC_block_wb = load_workbook('Block IM Jul-Dec 2024.xlsx', data_only=True, keep_vba=True)
BMC_EARS_wb = xw.Book('Internal Medicine EARs AY25 BMC.xlsm')
MONTH_NUM = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
irrelevant_sheets = ["HPT_List", "D", "CODES", "FINAL RECONCILIATION", "EAR_OVERVIEW"]


def populate_block_BMC(data_wb_filename, EARS_wb_filename):
    """Given a workbook of BMC BLOCK trainee data, populate the corresponding cells in the EARS workbook
    For the BMC Jul-Dec 2024 workbook, the excel sheet starts with row 3466
    Note:  the trainee data is a BLOCK spreadsheet, full day shifts, remember to check both boxes """

    # data_only=True is needed because each cell contains a formula, we want the value displayed not the formula
    # keep_vba=True is necessary to prevent macros from being stripped after openpyxl access workbook
    BMC_block_wb = xw.Book(data_wb_filename) # read
    BMC_EARS = xw.Book(EARS_wb_filename) # write

    # get first worksheet name from the workbook
    full_sheet = BMC_block_wb.sheets[0]

    row_count = 2 # this variable serves as a counter to keep track of current row
    print(full_sheet.cells(row_count, 6).value)
    # iterating down the BMC block spreadsheet until the row is empty (meaning no more inputs)
    while full_sheet.cells(row_count, 1).value != None: 
        # get the full name following the last_name, full_name format
        last_name = full_sheet.cells(row_count, 1).value
        first_name = full_sheet.cells(row_count, 2).value
        full_name = last_name + ", " + first_name

        # get the start_dates and end_dates
        start_date = full_sheet.cells(row_count, 9).value
        end_date = full_sheet.cells(row_count, 10).value

        # get an array of the all the datetime objects in between start and end
        dates_between = get_dates_between(start_date, end_date)

        for date in dates_between:
            # for each date, find the corresponding sheet and cell within the
            year = date.year
            month = date.month
            day = date.day
            EARS_sheet = get_sheet(BMC_EARS, year, month, day)

            if EARS_sheet == None:
                print('No sheet found for', full_name, 'on', date)
                continue
            
            cell_1 = get_cell(day, "AM", full_name)
            cell_2 = get_cell(day, "PM", full_name)
            
            if cell_1 != None:
                rotation_type = get_rotation_type(full_sheet, row_count)
                set_cell(cell_1, EARS_sheet, BMC_block_wb, EARS_wb_filename, rotation_type)
            else:
                print(full_name, 'not found in EARS spreadsheet')

            if cell_2 != None:
                rotation_type = get_rotation_type(full_sheet, row_count)
                set_cell(cell_2, EARS_sheet, BMC_block_wb, EARS_wb_filename, rotation_type)
            else:
                print(full_name, 'not found in EARS spreadsheet')
        row_count += 1
            


def get_dates_between(start_date, end_date):
    """Given two datetime objects, return an array of datetime objects 
    between the two dates (inclusive)"""
    dates = []
    current_date = start_date
    while current_date <= end_date:
        dates.append(current_date)
        current_date += timedelta(days=1)
    return dates

# def get_relevant_sheets(EARS_workbook):
#     """Given a EARS workbook, return an array of only the releveant sheet objects
#     irrelevant sheets are: HPT_List, D, CODES, and FINAL RECONCILIATION, and EAR_OVERVIEW"""
#     relevant_sheets = []
#     for sheet in EARS_workbook:
#           if sheet.title not in irrelevant_sheets:
#             relevant_sheets.append(sheet)
#     return relevant_sheets

def get_sheet(EARS_wb, year, month, day):
    """Given EARS workbook, year, month, and day, return the correct worksheet object 
    that corresponds with the given date"""
    given_date = datetime(year, month, day)
    for sheet in EARS_wb.sheets:
        if sheet.name not in irrelevant_sheets: 
            # check year associated with sheet
            sheet_year = int(sheet["C9"].value)
            if sheet_year == year: # Look in sheets with correct year
                sheet_month = MONTH_NUM[sheet["C8"].value]
                sheet_start_date = datetime(sheet_year, sheet_month, 1)
                sheet_end_date = sheet["G4"].value
                if sheet_start_date <= given_date and given_date <= sheet_end_date: # Find sheet with correct date range ex: 7/1 - 7/31
                    return sheet
    # if no sheet found, return none
    return None

def make_names_dict(EARS_wb):
    """Given the EARS workbook, return a dictionary that holds all name to row mappings"""
    names_dict = {}
    # This function assumes that all names are consistent in every sheet, so instead of generating a dictionary
    # for each worksheet, we just get the first one from the workbook and create a 1D dict of name : coordinatr
    sheet = EARS_wb.sheets[5] # EAR_Jul_24
    curr_row = 13
    while sheet.cells(curr_row, 2).value != None:
        name = sheet.cells(curr_row, 2).value
        names_dict[name] = "B" + str(curr_row)
        curr_row+=2
    return names_dict

def get_cell(day, shift, name):
    """Given the EARS worksheet object, day, first_shift status, and name, 
    return the corresponding EARS sheet cell coordinate"""
    # check that the name is actually on the spreadsheet

    # if the name passed into this function has a title ex: "Vergara Greeno, Rebeca (DGM)"
    name = name.split(' (')[0]
    if name not in BMC_name_mappings:
        return None # name not found on spread sheet
    str_cor = ""
    if 73 + day - 1 > 90:
        str_cor += "A"
        str_cor += chr(73 + day - 1 - 90 + 64)
    else:
        str_cor += chr(73 + day - 1)
    
    if shift == 'AM':
        row_num = BMC_name_mappings[name][1:] # we only want the row number, ignore the first letter
        str_cor += row_num
    elif shift == 'PM':
        row_num = int(BMC_name_mappings[name][1:]) + 1
        str_cor += str(row_num)

    return str_cor

def set_cell(cor, sheet, workbook, EARS_wb_filepath, rotation_type):
    """Given the EARS sheet, the EARS coordinate, the EARS workbook object, and the EARS workbook file path
    set the value within that sheet to present."""
    sheet.range(cor).value = rotation_type
    print("set", cor, "in file:", sheet)
    # save changes to the file after all changes have been made
    workbook.save()


def get_rotation_type(BMC_block_sheet, BMC_block_row):
    """Given the BMC block spreadsheet and the current row, retrieve and return a 
    string representing the rotation type (P or PTO)"""
    rotation_defin = BMC_block_sheet.cells(BMC_block_row, 6).value # column 6 contains rotation definition
    if "vacation" in rotation_defin.lower():
        return "PTO" # Paid Time Off
    else:
        return "P" # Present

BMC_name_mappings = make_names_dict(BMC_EARS_wb)
populate_block_BMC('Block IM Jul-Dec 2024.xlsx', 'Internal Medicine EARS AY25 BMC.xlsm')