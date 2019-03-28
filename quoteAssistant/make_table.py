import win32com.client as win32
from pathlib import Path

def make_table(output_file):
    try:
        excel = win32.dynamic.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(output_file)
        excel.Visible = False
    except Exception as e:
        print(e)
        wb = None

    op_entry_sheet = wb.Worksheets['operation_entry']

    wb.Worksheets.Add().Name = 'operations_pivot'

    operations_pivot = wb.Worksheets['operations_pivot']

    pivot_cache = wb.PivotCaches().Create(SourceType=win32.constants.xlDatabase,
                                                      SourceData=op_entry_sheet.UsedRange)

    pivot_table = pivot_cache.CreatePivotTable(TableDestination=operations_pivot.Range("A1"))

    c = win32.constants

    subassembly = pivot_field('subassembly',pivot_table,'Page')
    op_id = pivot_field('op_id',pivot_table,'Row')
    row = pivot_field('row',pivot_table,'Data','Average','0')
    op_id.name.AutoSort(c.xlAscending,row.name)

    qty = pivot_field('qty',pivot_table,number_format='0')
    setup = pivot_field('setup',pivot_table,number_format='0.00')
    labor = pivot_field('labor',pivot_table,number_format='0.00')
    pivot_table.CalculatedFields().Add('prod_std', "=Labor*60/Qty")

    prod_std = pivot_field('prod_std',pivot_table,number_format='0.000')

    #TODO: materials table
    #TODO: subcontracts table

class pivot_field:

    def __init__(self,
                 name,
                 pt,
                 orientation='data',
                 function=None,
                 number_format=None,
                 position=None):

        c = win32.constants
        self.name = pt.PivotFields(name)

        if str('Data') in orientation.title():
            self.name.Orientation = c.xlDataField
        elif str('Row') in orientation.title():
            self.name.Orientation = c.xlRowField
        else:
            self.name.Orientation = c.xlPageField
            self.name.EnableMultiplePageItems = True

        if position is not None:
            self.name.Position = position

        if function is not None:
            if str('Sum') in function.title():
                self.name.Function = c.xlSum
            elif str('Average') in function.title():
                self.name.Function = c.xlAverage
            else:
                self.name.Function = c.xlCount

        if number_format is not None:
            self.name.NumberFormat = number_format
