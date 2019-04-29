import win32com.client as win32

def load_spreadsheet(file=None):
    try:
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
        except:
            excel = win32.dynamic.Dispatch('Excel.Application')
            
        excel.DisplayAlerts = False

        if file is not None:
            wb = excel.Workbooks.Open(file)
        else:
            wb = excel.Workbooks.Add()

    except Exception as e:
        print(e)
        wb.Close(True)
        quit()

    finally:
        excel.Visible = False
        return(wb)
