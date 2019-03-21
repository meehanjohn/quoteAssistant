import win32com.client as win32

def load_spreadsheet(file=None):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        if file is not None:
            wb = excel.Workbooks.Open(file)
        else:
            wb = excel.Workbooks.Add()
        excel.Visible = False
    except Exception as e:
        print(e)
        wb = None
    finally:
        return(wb)
