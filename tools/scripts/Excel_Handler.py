import win32com.client as win32
import os

from mops import Mops
from mops import MopsConstants as MC
from BillingLetterGenerator import BL_Plugins
from BillingLetterGenerator import BL_Constants as BLc

def initialize_excel (visible = False):
    # This method starts and returns for use an instance 
    #    of Microsoft Excel to open the validation document
    #
    # Inputs:
    #    visible - Boolean, default False
    #        set to true in order to watch MW while program
    #        is working
    #
    # Outputs:
    #    excel - the MS Excel instance
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = visible
    return excel
    
def open_workbook (excel, workbook_path = None):
    # This method opens the excel workbook. If no path is 
    #    provided, a blank wb is returned
    # 
    # Inputs:
    #    excel     - MS Excel Instance
    #    Workbook_path 
    #            - Path to the workbook, default None
    #
    # Outputs:
    #    workbook - new or user specified workbook
    
    if workbook_path:
        return excel.Workbooks.Open (workbook_path)
    else:
        return excel.Workbooks.Add()
    
def get_attr_mapping (data_sheet, used_rows):
    # This method gets the attribute keys to be used in the 
    #    company data. They are the values in the first column
    #    of the vendor data excel file and specify where all the
    #     information needed is. These values should line up with
    #     those in the VENDOR ATTRIBUTE TAGS section of BL_Constants
    #
    # Inputs:
    #    data_sheet
    #            - Excel workbook sheet to be parsed
    #    used_rows 
    #            - max rows with data in data_sheet
    #
    # Outputs:
    #     attrs_mapping_dict = {
    #                attribute_key_1 : [List of relevant rows],
    #                attribute_key_2 : [...],
    #                ...
    #            }
    
    # Initialize the dictionary
    attrs_mapping_dict = {}
    
    # Bool to skip empty rows at beginning if any    
    allow_none = False
    
    # Initialize key (for scope)
    key = None
    
    # scroll through all used rows
    for row in range(1, used_rows+1):
        # get the cell value
        cell_val = data_sheet.Range('A%d'%row).Value
        
        # if there is a cell value, make it an attribute key
        # and add the current row as first in the relevant
        # list
        if not cell_val == None:
            key = cell_val
            allow_none = True
            attrs_mapping_dict[key] = [row]
            
        # if no cell value and not initial blank cells, 
        # add row to the attributes list of relevant rows
        elif allow_none:
            attrs_mapping_dict[key].append(row)
    
    return attrs_mapping_dict    
    
def get_col_letters (col):
    # This method takes the number of the column being processed and 
    #     returns a CHAR column identifier for MS Excel to use
    #
    # Inputs:
    #    col - integer of column
    #
    # Outputs:
    #    col_letter - the letter that corrosponds to the column number
    
    col_letter = None
    # if the col is in the first (one char length) set of columns
    # simply complete it's ascii value and convert to char
    if col <= 26:
        col_letter = chr (col + 64)
    # if the value is past the first set of columns, break col 
    # number by base 26 and recusively call self until a group of 
    # characters have been returned representing the excel column
    else:
        div = int(col/26)
        mod = int(col%26)
        
        # if there is no remainder the last character will be 'Z' as
        # its value is 26 in a base 26 system
        if mod == 0:
            div -= 1
            mod = 26
        
        letter_one = get_col_letters (div)
        letter_two = get_col_letters (mod)
        
        col_letter = letter_one + letter_two
        
    return col_letter        
    
def get_excel_data_by_company (data_sheet):
    # excel_data_by_company = {
    #        'Company1' : {
    #                BLc.VENDOR_PNAME_TAG    : value or [values,],
    #                BLc.TO_TAG                 : value or [values,],
    #                BLc.CC_TAG                 : value or [values,],
    #                BLc.ATTACH_TAG             : value or [values,],
    #                BLc.EMAIL_MESSAGE_TAG    : value or [values,],
    #                BLc.LETTER_MESSAGE_TAG     : value or [values,],
    #                BLc.GREET_TAG             : value or [values,],
    #                BLc.CLOSING_TAG         : value or [values,],
    #                BLc.NOTES_TAG             : value or [values,],
    #                BLc.CONTACT_ADDRESS_TAG : value or [values,]
    #            },
    #        'Company2' : {...},
    #        ...
    #        }

    # Get the used area of the sheet
    used_range = data_sheet.UsedRange
    used_rows = used_range.Rows.Count
    used_cols = used_range.Columns.Count
    
    # get attributes designated rows as a dict so that they can
    # be selected individually
    attrs_mapping_dict = get_attr_mapping (data_sheet, used_rows)
    
    company_data = {}
    company_name = None
    
    # start on col 2 because col 1 is row labels
    # add one to used_cols because range function is
    # noninclusive
    for col in range (2, used_cols + 1):
        col_letter = get_col_letters (col)
        
        # 
        for key in attrs_mapping_dict.keys ():
            # Add the data from the relevant rows to the vendor's attribute
            for row in attrs_mapping_dict[key]:
                cell_value = data_sheet.Range('%s%d' % (col_letter, row)).Value
                if cell_value == None:
                    continue
                
                # if the data is the vendor name, then add it as a new
                # key in the company 
                if key == BLc.VENDOR_TAG:
                    company_name = cell_value
                    company_data [company_name] = {}
                    
                # if the data is not the vendor, take it as an attribute 
                # and place it in the nested dictionary
                else:
                    if key in company_data[company_name].keys():
                        if not isinstance (company_data[company_name][key], list):
                            company_data[company_name][key] = [
                                    company_data[company_name][key]
                                ]
                        company_data[company_name][key].append(cell_value)
                    else:
                        company_data[company_name][key] = cell_value

    return company_data

def get_rel_data_from_ws (ws, data_dict):
    # This method takes a dictionary of keys and 
    #    finds their corrosponding values in the 
    #    pointed worksheet
    #
    # Inputs:
    #    ws - MS Excel Worksheet with relevant data
    #    data_dict - empty dict w keys to match data
    #
    # Outputs:
    #    data_dict - returns filled with values for keys

    # Get used row count
    used_range = ws.UsedRange
    rows = used_range.Rows.Count
    
    for row in range (1, rows + 1):
    
        tag_cell = ws.Range('A%d' % row).Value
        if isinstance(tag_cell, str):
            tag_cell = tag_cell.strip()
        
        # If cell matches a key, get the value next to it
        if tag_cell in data_dict.keys():
            
            data_dict [tag_cell] = ws.Range ('B%d' % row).Value
    
    return data_dict    
    
def get_run_variables_from_excel (vars_sheet):
    # This method initializes an empty data dictionary to 
    #    be filled with the settings for this run
    #
    # Inputs:
    #    vars_sheet - MS Excel Worksheet with settings information
    #
    # Outputs:
    #    run_vars_data - dictionary of run settings
    
    # Init dict
    run_vars_data = {
        BLc.YEAR_TAG : None,
        BLc.QUARTER_TAG : None,
        BLc.SUBJECT_TAG : None,
        BLc.DATABASE_TAG : None,
        BLc.NETWORK_DIR_TAG : None,
        BLc.M_NAME : None,
        BLc.M_TITLE : None
        }
    
    # Fill dict
    run_vars_data = get_rel_data_from_ws (vars_sheet, run_vars_data)
    
    return run_vars_data
    
def get_query_from_excel (query_sheet):
    # This method returns two sql queries responsible for vendor data
    #     and attachment information
    #
    # Inputs:
    #    query_sheet - MS Excel Worksheet with sql query information
    #
    # Outputs:
    #    main_query - the query for the vendor data
    #    files_query - the query to get the file list information
    
    query_data = {
        BLc.FILE_LIST_QUERY : None,
    }
    queries_data = get_rel_data_from_ws (query_sheet, query_data)
    
    files_query = query_data[BLc.FILE_LIST_QUERY]
    
    return files_query
    
def get_pivot_categories (piv_ops_sheet):
    # This method creates a dictionary that structures the pivot
    #    table
    #
    # Inputs:
    #    piv_ops_sheet - MS Excel Worksheet with pivot options
    #
    # Outputs:
    #    pivot_categories - Dictionary of pivot table formatting
    #                            information
    
    # Get used row count
    used_range = piv_ops_sheet.UsedRange
    used_rows = used_range.Rows.Count
    
    pivot_categories = {}
    
    # Start on row two as first row is headers
    for row in range (2, used_rows + 1):
        # Extract data
        category     = piv_ops_sheet.Range ('A%d' % row).Value
        xl_range     = piv_ops_sheet.Range ('B%d' % row).Value
        hierarchy     = piv_ops_sheet.Range ('C%d' % row).Value
        type_tag     = piv_ops_sheet.Range ('D%d' % row).Value
    
        # Create a dict entry for the category
        pivot_categories [category] = (xl_range, int(hierarchy), type_tag)
    
    return pivot_categories

def add_data_to_ws (ws, data, location, style = None):
    # This method appends data to the worksheet
    #
    # Inputs:
    #    ws - MS Excel Worksheet to be written to
    #    data - the information to go in the cell
    #    location - the cell location (row, col)
    #    style - if there is styling, pass a constant to 
    #                catch the case. Default None
    #
    # Outputs:
    #    None - writes to Worksheet
    
    # get excel format cell location
    row, col = location
    cell_loc = ('%s%d' % (get_col_letters(col), row))
    
    # Input data to cell
    ws.Range(cell_loc).Value = data
    
    # Add style
    if style:
        if style == BLc.BOLD:
            ws.Range (cell_loc).Font.Bold = True
        elif style == BLc.ALIGN_LEFT:
            ws.Range (cell_loc).HorizontalAlignment = win32.constants.xlHAlignLeft
        elif style == BLc.WRAP_MERGE:
            used_cols = ws.UsedRange.Columns.Count
            if used_cols < 4:
                used_cols += 1
            
            merge_range = ws.Range(
                                'A%d:%s%d' % (
                                    row, 
                                    get_col_letters (used_cols),
                                    row
                                )
                            )
            merge_range.MergeCells = True
            
            ws.Range (cell_loc).WrapText = True
            #autofit(ws)
            
    return
    
def get_files_list (source, files_query, company):
    # This method gets a fles list from a sql database
    #
    # Inputs:
    #    source - source string for DB
    #    files_query - the query to execute and grab file names
    #    company - the current vendor
    #
    # Outputs:
    #    files - list of filenames to be attached to email
    #    or None, if none
    
    # Connect to DB
    conn = win32.Dispatch (r'ADODB.Connection')
    rs = win32.Dispatch (r'ADODB.Recordset')
    
    # Place vendor name into query
    files_query = files_query.replace (
                    BLc.QUERY_VENDOR_TAG,
                    company)
    
    # Modify source string (does not need first OLEDB tag_cell
    #    when not read directly into PivotCaches
    source = ';'.join(source.split (';')[1:]).replace('=.','=PDXVMDB11').replace('Claims','TAI')
    
    # Open connection
    conn.Open (source)
    
    # Perform Query
    rs.Open(files_query, conn, 1, 3)
    
    # If data, return
    if not rs.RecordCount == 0:
        files = rs.GetRows ()[0]
        conn.Close()
        return files
    conn.Close()
    return None
    
def format_date (ws, cell):
    # Formats the cell to the date format provided
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #    cell - the cell to be formatted
    #
    # Outputs:
    #    None
    
    row, col = cell
    c = ws.Cells(row, col)
    
    ws.Range(c,c).NumberFormat = BLc.NUMBER_FORMAT
    return
    
def autofit (ws):
    # This method resizes the cells on the worksheet to fit 
    # more neatly
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #
    # Outputs:
    #    None

    # Auto-Size 
    Mops.autofit_worksheet(ws)
    
    # Merged and Centered cells will not adjust with autofit,
    #    manually set rowheight
    ws.Rows (BLc.PRE_TABLE_TEXT_ROW).RowHeight = \
        BLc.PRE_TABLE_ROW_HEIGHT
    ws.Rows (BLc.ADDRESS_ROW).RowHeight = \
        BLc.ADDRESS_ROW_HEIGHT
    
    # Auto-Size Pivot Table Area
    ws.PivotTables(1).TableRange1.Columns.AutoFit ()
    ws.Columns.AutoFit ()
    
    return
    
def get_next_available_row (ws):
    # Returns the next available row on the worksheet
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #
    # Outputs:
    #    next_available - int, next available row in ws
    
    current_length = ws.UsedRange.Rows.Count
    
    # Iterate by 2 to allow some space
    next_available = current_length + 2
    
    return next_available
    
def get_next_available_cell (ws):
    # Returns the first cell of the next available row
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #
    # Outputs:
    #    CELL INSTANCE - next available cell

    return (get_next_available_row (ws), 1)

def create_range_from_cell (ws, cell):
    # Turns a cell into a range object
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #    cell - the cell to be converted
    #
    # Outputs:
    #    RANGE OBJECT - a range generated on the cell location
    row, col = cell
    print (type(ws))
    print ('Cell 1,1 ', ws.Cells(1,1))
    print ('Cell 1,2 ', ws.Cells(1,2))
    print ('Cell 1,3 ', ws.Cells(1,3))
    print ('Cell 1,4 ', ws.Cells(1,4))
    print ('Cell 1,5 ', ws.Cells(1,5))
    print ('Cell 1,6 ', ws.Cells(1,6))
    print ('Cell 1,7 ', ws.Cells(1,7))
    print ('Cell 1,8 ', ws.Cells(1,8))
    print ('Cell 1,9 ', ws.Cells(1,9))
    print ('Cell 1,10 ', ws.Cells(1,10))
    print ('Cell 1,11 ', ws.Cells(1,11))
    
    ws.Cells(row, col)
    print("--- " ,ws.Range("A11"))
    c = ws.Cells (row, col)
    return ws.Range (c, c)

def update_treaties (run_variables):
    # This method invokes SQL to update the treaties in the table before
    # making specifig queries that require this information to have been
    # updated.
    # 
    # Inputs:
    #    run_variables - the user specific options
    #
    # Outputs:
    #     None
    
    # Get year and quarter info
    year = int (run_variables [BLc.YEAR_TAG])
    quarter = int (run_variables [BLc.QUARTER_TAG])
    
    # Create a year/month string that is used in the database
    yearmo = '%d02%d' % (year, quarter*3)

    # Get the database source string
    source_raw = run_variables [BLc.DATABASE_TAG]
    source = ';'.join (source_raw.split (';') [1:])
    
    # format the query so it is complete
    query = BLc.TREATIES_QUERY.format (yearmonth = yearmo)
    
    # Query the database
    conn = win32.gencache.EnsureDispatch(r'ADODB.Connection')
    rs = win32.gencache.EnsureDispatch (r'ADODB.Recordset')
    conn.Open (source)
    rs.Open (query, conn, 1, 3)
    conn.Close ()
    
    return
    
def style_pivot (pt, pivot_categories, company_data, title):
    # This method assigns the pivot fields to their locations and 
    # whether or not to display them. Also toggles some pivot table
    # options.
    #
    # Inputs:
    #    pt - PivotTable instance
    #    pivot_categories - user dictionary with pivot field preferences
    #    company_data - dictionary with vendor data
    #    title - the title for the pivot table
    #
    # Outputs:
    #     None
    
    # get win32 constants
    win32c = win32.constants
    
    # for each user defined pivot category...
    for category in pivot_categories.keys():
        xl_range, hierarchy, type_tag = \
            pivot_categories[category]
        
        # ...only display if requested
        if xl_range == BLc.PIVOT_NONE:
            continue
        
        # ...only display if correct type (if applicable)
        elif (
                    (
                    type_tag == BLc.PAYEE
                    and 
                    company_data [BLc.VENDOR_TYPE] == BLc.PAYER
                )
                or     (
                    type_tag == BLc.PAYER
                    and 
                    company_data [BLc.VENDOR_TYPE] == BLc.PAYEE
                )
            ):
            continue
        
        # ...add category as a pivot row
        elif xl_range == BLc.PIVOT_ROW:
            pt.PivotFields (category).Orientation = win32c.xlRowField
            pt.PivotFields (category).Position = hierarchy
            
        # ...add category as a pivot column
        elif xl_range == BLc.PIVOT_COL:
            pt.PivotFields (category).Orientation = win32c.xlColumnField
            pt.PivotFields (category).Position = hierarchy
            # disable column subtotals
            pt.PivotFields (category).Subtotals = BLc.SUBTOTAL_FALSE
            
        # ...add category as a pivot data field
        elif xl_range == BLc.PIVOT_PG:
            pt.AddDataField (pt.PivotFields (category))
            pt.PivotFields ('sum of %s' % category).NumberFormat = \
                BLc.MONEY_FORMAT
    
    # Order by Premiums, Claims, Claims Interest 
    if (BLc.PREMIUMS and BLc.CLAIMS and BLc.CLAIMS_INTEREST) in \
        [piv_item.Name for piv_item in pt.PivotFields(BLc.SECTION).PivotItems()]:
        pt.PivotFields(BLc.SECTION).PivotItems(BLc.PREMIUMS).Position = 1
        pt.PivotFields(BLc.SECTION).PivotItems(BLc.CLAIMS).Position = 2
        pt.PivotFields(BLc.SECTION).PivotItems(BLc.CLAIMS_INTEREST).Position = 3
    else:
        pt.PivotFields(BLc.SECTION).AutoSort (win32c.xlDescending, BLc.SECTION)
    
    
    # Toggle pivot table options
    pt.TableStyle2 = BLc.TABLE_STYLE
    pt.DisplayFieldCaptions = False
    pt.PrintTitles = False
    pt.ShowDrillIndicators = False
    pt.GrandTotalName = BLc.TOTAL_OWED
    pt.Name = title
    pt.RowGrand = False
    
    return
    
def generate_pivot (ws, query, source, range_, title, company, company_data, pivot_categories, run_variables):
    # This method generates the pivot table for the letter
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #    query - the query to be used to get the vendor data
    #    source - the database source string
    #    range_ - the area to insert the pivot table 
    #                (exta _ on name for convention protection)
    #    title - the pivot table's title
    #    company - the current vendor
    #    company_data - the dict of information for the vendor
    #    pivot_categories - user dictionary with pivot field references
    #    run_variables - user dictionary with run settings
    #
    # Outputs:
    #    None
    
    # Get win32 constants
    win32c = win32.constants
    
    # Init q for scope
    q = ''
    
    # If a corner case query is needed, perform that action
    if company_data [BLc.CORNER_CASE_TAG] in BLc.QUERY_SPECIFIC_CORNER_CASE_SET:
        q = BL_Plugins.specific_query_handler (company_data, run_variables)
    
    else:
        # Perform standard query
        q = query.format(
            yr=int(run_variables[BLc.YEAR_TAG]),
            mo=(int(run_variables[BLc.QUARTER_TAG]*3)),
            dy=Mops.month_len(int(run_variables[BLc.QUARTER_TAG]*3)))
        q = q.replace (BLc.QUERY_VENDOR_TAG, company_data[BLc.QUERY_TAG])
    
        # Replace payer/payee tag with relevant vendor type
        if company_data[BLc.VENDOR_TYPE] == BLc.PAYEE:
            q = q.replace (BLc.PAYEE_PAYER, BLc.PAYEE)
        elif company_data [BLc.VENDOR_TYPE] == BLc.PAYER:
            q = q.replace (BLc.PAYEE_PAYER, BLc.PAYER)
    
    # Make the pivot table from the info returned by the query
    pc = ws.Parent.PivotCaches().Add(win32.constants.xlExternal)
    
    pc.Connection = source
    pc.CommandType = win32.constants.xlCmdDefault
    pc.CommandText = q
    
    print (range_, title)
    
    pt = pc.CreatePivotTable (range_, title)
    
    # add style to the pivot
    style_pivot (pt, pivot_categories, company_data, title)
    
    # check for additional steps corner cases
    if company_data [BLc.CORNER_CASE_TAG] in BLc.EXTRA_STEPS_CORNER_CASES:
        BL_Plugins.extra_steps_handler(company_data [BLc.CORNER_CASE_TAG], ws, pt)
        
    return
    
def wrap_pivot (ws):
    # This method outlines the pivot table with a border
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #
    # Outputs:
    #    None
    
    # Get win32 constants
    win32c = win32.constants

    # Get used range
    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count
    
    # Add one to account for hidden 'sum of amounts' row
    top_row = BLc.PIVOT_ROW_LOC + 1
    
    # Get last used column
    top_col = get_col_letters (1)
    bot_col = get_col_letters (cols)
    
    # Set ranges
    top_range = ws.Range('%s%d:%s%d' % (top_col, top_row, bot_col, top_row))
    right_range = ws.Range ('%s%d:%s%d' % (bot_col, top_row, bot_col, rows))
    bottom_range = ws.Range ('%s%d:%s%d' % (top_col, rows, bot_col, rows))
    left_range = ws.Range ('%s%d:%s%d' % (top_col, top_row, top_col, rows))
    
    # Outlone the pivot table
    top_range.Borders(win32c.xlEdgeTop).LineStyle = win32c.xlContinuous
    right_range.Borders (win32c.xlEdgeRight).LineStyle = win32c.xlContinuous
    bottom_range.Borders (win32c.xlEdgeBottom).LineStyle = win32c.xlContinuous
    left_range.Borders (win32c.xlEdgeLeft).LineStyle = win32c.xlContinuous
    
    return
    
def hide_row (ws, row):
    ws.Rows(row).EntireRow.Hidden = True
    return
    
def get_total_width (ws, cols):
    width = 0
    for col in range (1, cols+1):
        cl = get_col_letters (col)
        width += \
            ws.Range('%s1' % cl).ColumnWidth
    return width

def get_subj_len (ws):
    row = BLc.SUBJECT_LOCATION [0]
    ws.Columns.AutoFit ()
    return ws.Range('A%d' % row).ColumnWidth    
    
def get_print_area (ws, hide_rows):
    # This method gets the area of the worksheet to be formatted
    # into a pdf document
    #
    # Inputs:
    #    ws - the MS Excel Worksheet
    #
    # Outputs:
    #    print_area - Range value representing the document space
    
    num_extra = 0
    row = BLc.SUBJECT_LOCATION [0]
    
    # Get the used range
    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count
    
    # Get the length of the subject 
    subj_len = get_subj_len (ws)
    
    # Merge the subject cell so that it doesn't affect
    # pivot sizing on AutoFit
    merge_range = ws.Range(
        'A%d:%s%d' % (
            row, 
            get_col_letters (ws.UsedRange.Columns.Count),
            row
        )
    )
    merge_range.MergeCells = True
    
    # Autofit the worksheet
    autofit (ws)
    
    # Get worksheet width
    width = get_total_width (ws, cols)
    
    # add extra width if necessary to ensure no cutoff
    if subj_len > width:
        diff = subj_len - width 
        num_extra = int(diff / 8)
        if diff % 8 > 0:
            num_extra += 1
    
    # Hide any pivot table top rows
    if hide_rows:
        for row in hide_rows:
            hide_row(ws, row)
    else:
        hide_row (ws, BLc.PIVOT_ROW_LOC)
    col_letter = get_col_letters (cols + num_extra)
    
    # Get range to used area on worksheet
    range_begin = MC.TOP_LEFT
    range_end = '%s%d' % (col_letter, rows)
    
    print_area = range_begin + ':' + range_end
    
    return print_area

def print_pdf (wb, path, hide_rows):
    # This method converts the excel to a pdf
    #
    # Inputs:
    #    wb - the MS Excel Workbook
    #    path - the pdf outfile path
    #
    # Outputs:
    #    None
    
    ws = wb.Worksheets (1)
    
    # get the print area from the worksheet
    print_area = get_print_area (ws, hide_rows)
    
    # print options
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
    
    # export as PDF
    ws.ExportAsFixedFormat(0, path)
    
    return

def close_excel (xl, verbose):
    # This method stops the excel process
    print (BLc.STAT_EXIT_EXCEL)
    xl.Application.Quit ()
    obliterate (verbose)
    return

def obliterate (verbose=False):
    os.system('Taskkill /IM EXCEL.exe /F')
    if verbose:
        print (BLc.SINGLE_NEWLINE + BLc.QUOTH_THE_BONES)
    return