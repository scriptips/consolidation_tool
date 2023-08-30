import win32com.client as win32
import time
import pywintypes

def consolidate_wage_sheets():
    '''Last update 2023-08-29. Working version. .exe file created with pyinstaller.'''

    # Connecting to the open excel file
    xl = win32.GetActiveObject("Excel.Application")
    wb = xl.ActiveWorkbook
    wb_name = wb.Name
    wb.AutoSaveOn = False

    print('\nINFO: >>> Consolidation In Progress, Please Wait...\n')

    # Set calculation mode to Manual
    xl.Calculation = -4135
    
    # capturing the year and month from the file name
    year = ''.join(['20',(wb_name.split(' ')[0]).split('.')[1]])
    month = (wb_name.split(' ')[0]).split('.')[0]
    year_dash_month = '-'.join([year,month])

    dati_sh = wb.Worksheets('DATI') 
    non_touchable = ['DARBINIEKI', 'LIKMES', 'LAIKI', 'VALSTIS', 'DATI', 'PIVOT']
    curr_file_sht_list = [x.Name for x in wb.Worksheets]
    curr_file_sht_list = [x for x in curr_file_sht_list if x not in non_touchable]

    spot = dati_sh.Cells(dati_sh.UsedRange.Rows.Count, 6).Row
    dati_sh_rng = dati_sh.Range(f'A1:A{spot}')

    x = []
    for year_month_cell in dati_sh_rng:
        if year_month_cell.Value == year_dash_month:
            x.append(year_month_cell)

    for date_cell in dati_sh_rng:
        if str(date_cell) == str(year_dash_month):
            date_cell.EntireRow.ClearContents()

    # DATI sheet consolidation
    for rng in curr_file_sht_list:
        # Column F chosen to always have a criteria filled row to take the sum from, Eur,
        # as sometimes activities, networks etc. can be missing...
        filled_pr_cost_lines = [33 + n for n, x in enumerate(wb.Worksheets(rng).Range('F33:F43')) if x.Value]
        for row in filled_pr_cost_lines:
            # 'spot' defines the first empty row after the already filled region, to continue from there...
            # taking the sum from the criteria column F, Eur..
            spot = dati_sh.Cells(dati_sh.UsedRange.Rows.Count, 6).Row + 1
            for col in range(2, 7):
                dati_sh.Cells(spot, col + 1).Value = wb.Worksheets(rng).Cells(row, col).Value
                if dati_sh.Cells(spot, col + 1).Value:
                    dati_sh.Cells(spot, 1).Value = f'{wb.Worksheets(rng).Cells(2, 7).Value}-{str(wb.Worksheets(rng).Cells(3, 7).Value).zfill(2)}'
                    if wb.Worksheets(rng).Cells(2, 16).Value is None:
                        dati_sh.Cells(spot, 2).Value = wb.Worksheets(rng).Cells(3, 16).Value
                    else:  dati_sh.Cells(spot, 2).Value = wb.Worksheets(rng).Cells(3, 16).Value + ' ' + wb.Worksheets(rng).Cells(2, 16).Value

    # PIVOT sheet update
    extd_source_data_rng = f'DATI!$A$1:$G${str(spot)}'
    pivot_sh = wb.Worksheets('PIVOT')
    pivot_table = pivot_sh.PivotTables('PivotTable1')
    pivot_table.ChangePivotCache(wb.PivotCaches().Create(SourceType=1, SourceData=extd_source_data_rng))
    pivot_table.RefreshTable()

    pivot_field = pivot_table.PivotFields('Month')
    pivot_field.ClearAllFilters()

    try:
        pivot_field.CurrentPage = year_dash_month
    except pywintypes.com_error:
        print('\nINFO: >>> No data for the selected month! Selecting All...\n')
        pivot_field.CurrentPage = '(All)'
        time.sleep(4)

    # Little formatting..
    c_align = pivot_sh.Range("C:C")
    c_align.HorizontalAlignment = -4152
    d_align = pivot_sh.Range("D:D")
    d_align.HorizontalAlignment = -4152
    pivot_sh.Range("A1")
    dati_sh.Visible = False
    

    # Set calculation mode back to Automatic
    xl.Calculation = -4105
    # Trigger calculations
    
    wb.Save()

    print('\nINFO: >>> Consolidation Done Successfully! Exiting...\n')
    time.sleep(4)


if __name__ == '__main__':
    consolidate_wage_sheets()
    input('\nPress Enter to exit...')