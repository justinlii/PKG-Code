import xlwings as xw # must install pip before installing xlwings
import os
from datetime import datetime, timedelta

BMC_assign_wb = xw.Book('Assign Name Jul-Dec 2024.xlsx', data_only=True, keep_vba=True)
BMC_EARS_wb = xw.Book('Internal Medicine EARs AY25 BMC.xlsm')
MONTH_NUM = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
irrelevant_sheets = ["HPT_List", "D", "CODES", "FINAL RECONCILIATION", "EAR_OVERVIEW"]


def populate_block_BMC(data_wb_filename, EARS_wb_filename):
    """Given a workbook of BMC BLOCK trainee data, populate the corresponding cells in the EARS workbook
    For the BMC Jul-Dec 2024 workbook, the excel sheet starts with row 3466
    Note:  the trainee data is a BLOCK spreadsheet, full day shifts, remember to check both boxes """

    # data_only=True is needed because each cell contains a formula, we want the value displayed not the formula
    # keep_vba=True is necessary to prevent macros from being stripped after openpyxl access workbook

    exceptions = []

    BMC_assign_wb = xw.Book(data_wb_filename) # read
    BMC_EARS = xw.Book(EARS_wb_filename) # write

    # get second worksheet name from the workbook (it should be titled "Detail")
    full_sheet = BMC_assign_wb.sheets[1]

    row_count = 2 # this variable serves as a counter to keep track of current row
    # iterating down the BMC block spreadsheet until the row is empty (meaning no more inputs)
    while full_sheet.cells(row_count, 1).value != None: 
        # get the full name following the last_name, full_name format
        last_name = full_sheet.cells(row_count, 1).value
        first_name = full_sheet.cells(row_count, 2).value
        full_name = last_name + ", " + first_name
        # get the start_dates and end_dates
        start_time = get_start_time(full_sheet, row_count)
        end_time = get_end_time(full_sheet, row_count)

        year = start_time.year
        month = start_time.month
        day = start_time.day

        # find the spreadsheet associated with the current name, 
        EARS_sheet = get_sheet(BMC_EARS, year, month, day)

        # if no spreadsheet found, print out the entry and skip
        if EARS_sheet == None:
            exceptions.append(f"No sheet found for, {full_name}, on this date: {start_time}")
            continue
        
        # Shifts on assign spreadsheet typically span 4-24 hours and not days, so we don't use get_dates_between
        shift_length = get_shift_length(full_sheet, row_count)
        
        # if shift length is 0 hours, skip
        if shift_length == 0:
            continue

        shift_start = ""
        if start_time.hour < 12:
            shift_start = "AM"
        else:
            shift_start = "PM"
        
        cells = []
        if shift_length < 7.5 and shift_length > 0: # mark one shift

            
            cell_cor = get_cell(day, shift_start, full_name)
            print('shift under 8 hours', shift_length, cell_cor)
            cells.append(cell_cor)
        elif shift_length >= 8 and shift_length <= 24: # mark two shifts
            cells.append(get_cell(day, shift_start, full_name))
            if shift_start == "AM":
                cells.append(get_cell(day, "PM", full_name))
            else: #if the shift starts in PM and continues into AM, mark the AM shift on the day after
                next_day = (start_time + timedelta(days=1)).day
                cells.append(get_cell(next_day , "AM", full_name))
        else:
            exceptions.append(f"Exception found for {full_name}, shift starts: {start_time}")
    
        for cell in cells:
            if cell == None:
                exceptions.append(f"{full_name} not found in workbook")
            else:
                rotation_type = get_rotation_type(BMC_assign_wb, row_count)
                set_cell(cell, EARS_sheet, BMC_assign_wb, EARS_wb_filename, rotation_type)
        row_count+=1
            

def get_start_time(sheet, row_count):
    """Given the 'assign name' sheet and the current row,
    return the correct start datetime object to use"""
    projected_start = sheet.cells(row_count, 14).value
    actual_start = sheet.cells(row_count, 15).value

    # if actual_start holds a valid date that is different from projected_start
    if actual_start != None and actual_start != projected_start:
        return actual_start # return actual_start and use this
    # otherwise return projected_start
    return projected_start

def get_end_time(sheet, row_count):
    """Given the 'assign name' sheet and the current row,
    return the correct end datetime object to use"""
    projected_end = sheet.cells(row_count, 16).value
    actual_end = sheet.cells(row_count, 17).value

    # if actual_end holds a valid date that is different from projected_end
    if actual_end != None and actual_end != projected_end:
        return actual_end # use actual_end date
    # otherwise return projected_end
    return projected_end

def get_shift_length(sheet, BMC_assign_row):
    """Given the BMC Assign sheet object and the current row,
    return the length of the shift (the shift length is recorded under "Actual Hours" column)"""
    return sheet.range((BMC_assign_row, 11)).value

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
    workbook.save(EARS_wb_filepath)


def get_rotation_type(BMC_sheet, curr_row):
    """Given the BMC assign spreadsheet and the current row, retrive and return a 
    string representing the rotation type (P or PTO)"""
    rotation_defin = BMC_sheet.range(curr_row, 6).value # column 6 contains rotation definition
    if "vacation" in rotation_defin.lower():
        return "PTO" # Paid Time Off
    else:
        return "P" # Present
    
BMC_name_mappings = make_names_dict(BMC_EARS_wb)
populate_block_BMC('Assign Name Jul-Dec 2024.xlsx', 'Internal Medicine EARS AY25 BMC.xlsm')


