import gspread
import re

def main():
    gc = gspread.service_account('/Users/zokeez/GoogleDriveSecrets/bazilona-google-spreadsheets.json')

    # Test table - https://docs.google.com/spreadsheets/d/1CNqiJZ72dX486MwzsgOUHAS6qYnuMoPW7l5egH4stKU/edit#gid=0

    # Open table by name
    # sh = gc.open('Python to google spreadsheets')
    # print(sh.sheet1.get('A1'))

    # Open table by ID (which can be found in URL)
    # sh = gc.open_by_key('1CNqiJZ72dX486MwzsgOUHAS6qYnuMoPW7l5egH4stKU')
    # print(sh.sheet1.get('A1'))

    # Open table by URL
    # sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1CNqiJZ72dX486MwzsgOUHAS6qYnuMoPW7l5egH4stKU/edit#gid=0')
    # print(sh.sheet1.get('A1'))

    # Create empty spreadsheet
    # sh = gc.create('New table 3')
    # # and share it to gmail account (not the owner)
    # sh.share('bazilona@gmail.com', perm_type='user', role='writer')

    sh = gc.open('Python to google spreadsheets')

    # # get worksheet by index
    # worksheet = sh.get_worksheet(1)
    # print(worksheet.get('A1'))

    # get worksheet by name
    # worksheet = sh.worksheet("Third")
    # print(worksheet.get('A1'))

    # list of worksheets
    # worksheets = sh.worksheets()
    # print(worksheets)

    # create worksheet
    # worksheet = sh.add_worksheet("Sheet from code", rows = 20, cols = 20, index = 1)

    # Delete worksheet
    # sh.del_worksheet(sh.worksheet("Shhet from code"))

    # Get cell value
    # print(sh.sheet1.acell('A1').value)
    # print(sh.sheet1.cell(1,1).value)
    # get formula
    # print(sh.sheet1.acell('A5', value_render_option='FORMULA').value)
    # print(sh.sheet1.cell(row=5, col=1, value_render_option='FORMULA').value)

    # Get all values from the column
    # values_list = sh.sheet1.col_values(4)
    # print(values_list)

    # Get all values from the row
    # values_list = sh.sheet1.row_values(8)
    # print(values_list)

    # Get all values from the worksheet
    # list_of_lists = sh.sheet1.get_all_values()
    # print(list_of_lists)

    # Get all values from the worksheet as a dictionary
    # sheet = sh.worksheet("Second")
    # list_of_dicts = sheet.get_all_records()
    # print(list_of_dicts)

    # Find a cell
    # sheet = sh.worksheet("4")
    # cell = sheet.find("Adincube")
    # print(cell)
    # print(f'Found Adincube in {cell.row}:{cell.col}')
    #
    # sheet = sh.worksheet("4")
    # cell = sheet.find(re.compile("Adin[Cc]ube"))
    # print(cell)
    # print(f'Found Adin[Cc]ube in {cell.row}:{cell.col}')
    #
    # sheet = sh.worksheet("4")
    # amount_re = re.compile(r'(Black|Red) silver')
    # cell = sheet.find(amount_re)
    # print(f'Found (Black|Red) silver in {cell.row}:{cell.col}')

    # find all cells
    # sheet = sh.worksheet("4")
    # cells = sheet.findall(re.compile(r'(Black|Red) silver'))
    # print(f'Found (Black|Red) silver in {cells}')

    # Cell properties
    # cell = sh.sheet1.acell('A3')
    # print(cell.value)
    # print(cell.row)
    # print(cell.col)

    # update cell
    # sh.sheet1.update('D2', "New value")
    # sh.sheet1.update_cell(3, 4, "New value")
    # sh.sheet1.update_cell(3, 4, "")

    # update range of the cells
    # sh.sheet1.update('F1:F6', [[1],[2],[3],[4],[5],[6]])
    # sh.sheet1.update('G1:H1', [[1,2]])

    # format
    #sh.sheet1.format('A1', {'text_format': {'bold': True}})
    # sh.sheet1.format('F1:F6',
    #  {
    #     'backgroundColor': {
    #         "red": 0.0,
    #         "green": 0.9,
    #         "blue": 0.3
    #     },
    #     'verticalAlignment': 'BOTTOM',
    #     'horizontalAlignment': 'LEFT',
    #     'text_format': {
    #         'foregroundColor': {
    #             "red": 0.9,
    #             "green": 0.2,
    #             "blue": 0.2
    #         },
    #         'bold': True,
    #         'italic': True,
    #         'underline': True,
    #         'strikethrough': True,
    #         'fontSize': 12,
    #         'fontFamily': 'Verdana'
    #     }
    # })

    # Group columns for hiding
    # sh.worksheet('4').add_dimension_group_columns(3,4)

    # add columns on the right
    # sh.worksheet('4').add_cols(7)

    # add protections to specific cells and define the list of allowed editors
    # sh.worksheet('4').add_protected_range('A1:B2', editor_users_emails=["bazilona@gmail.com","my-test-account@mythic-origin-356115.iam.gserviceaccount.com"])

    # freeze the top of rows and cols to show them always
    # sh.worksheet('5').freeze(rows=1, cols=2)

    # add and get notes
    # sh.worksheet('5').insert_note('A1', "123")
    # print(sh.worksheet('5').get_note('A1'))

    # merge cells
    # sh.worksheet('5').merge_cells('E4:E8')
    # sh.worksheet('5').merge_cells('F4:G4')
    # sh.worksheet('5').merge_cells('H4:K7', 'MERGE_COLUMNS')
    # sh.worksheet('5').unmerge_cells('H4:K7')

    # auto size
    # sh.worksheet('5').columns_auto_resize(8,9)

    # column width
    # sheetId = sh.worksheet('5')._properties['sheetId']
    # body = {
    #     "requests": [
    #         {
    #             "updateDimensionProperties": {
    #                 "range": {
    #                     "sheetId": sheetId,
    #                     "dimension": "COLUMNS",
    #                     "startIndex": 0,
    #                     "endIndex": 1
    #                 },
    #                 "properties": {
    #                     "pixelSize": 200
    #                 },
    #                 "fields": "pixelSize"
    #             }
    #         }
    #     ]
    # }
    # sh.batch_update(body)

    print('Done')

if __name__ == '__main__':
    main()
