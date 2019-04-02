import win32com.client as win32

def load_spreadsheet(file=None):
    try:
        excel = win32.dynamic.Dispatch('Excel.Application')

        if file is not None:
            wb = excel.Workbooks.Open(file)
        else:
            wb = excel.Workbooks.Add()

    except Exception as e:
        print(e)
        return

    finally:
        excel.Visible = False
        return(wb)
