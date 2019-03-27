import win32com.client as win32
from pathlib import Path

def make_table(output_file):
    try:
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
        except AttributeError:
            f_loc = r'C:\Users\Amuneal\AppData\Local\Temp\gen_py'
            for f in Path(f_loc):
                Path.unlink(f)
            Path.rmdir(f_loc)
            xl = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(output_file)
        excel.Visible = False
    except Exception as e:
        print(e)
        wb = None
    finally:
        return(wb)

    op_entry_sheet = wb.Worksheets['operation_entry']

    wb.Worksheets.Add().Name = 'operations_pivot'

    operations_pivot = wb.Worksheets['operations_pivot']

    pivot_cache = wb.PivotCaches().Create(SourceType=win32.constants.xlDatabase,
                                                      SourceData=op_entry_sheet.UsedRange)

    pivot_table = pivot_cache.CreatePivotTable(TableDestination=operations_pivot.Range("A1"))

    pivot_field('subassembly','Page')
    pivot_field('op_id','Row')
    pivot_field('row','Average')
    op_id.AutoSort(win32c.xlAscending,row)
    pivot_field('qty',number_format='0')
    pivot_field('setup')
    pivot_field('labor')
    pivot_table.CalculatedFields().Add('prod_std', "=Labor*60/Qty")
    pivot_field('prod_std',number_format='0.000')

class pivot_field:

    def __init__(self,
                 name,
                 orientation='data',
                 position=None,
                 function='xlSum',
                 number_format='0.00'):

        c = win32.constants
        self.name = pivot_table.PivotFields(name)

        if str('Data') in orientation.title():
            self.name.Orientation = c.xlDataField
        elif str('Row') in orientation.title():
            self.name.Orientation = c.xlRowField
        else:
            self.name.Orientation = c.xlPageField
            self.name.EnableMultiplePageItems = True

        if position is not None:
            self.name.Position = position

        if str('Sum') in function.title():
            self.name.Function = 'xlSum'
        elif str('Average') in functin.title():
            self.name.Function = 'xlAverage'
        else:
            self.name.Function = 'xlCount'

        self.name.NumberFormat = number_format
