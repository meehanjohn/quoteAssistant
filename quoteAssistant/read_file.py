import win32com.client as win32
from quoteAssistant import get_stuff as get

def read_file(wb):
    # keys are used to locate where the useful information is in the quote
    # TODO add end key (typical row of values where relevant info ends)

    sheetnames = [sheet.Name for sheet in wb.Worksheets]
    materials = []
    operations = []
    subcontracts = []

    try:
        for sheet_index, sheet_name in enumerate(sheetnames,1):
            print(sheet_index, sheet_name)
            sheet_obj = wb.Worksheets[sheet_index]
            args = (sheet_index,sheet_name,sheet_obj)

            materials.extend(get.materials(*args))
            operations.extend(get.operations(*args))
            subcontracts.extend(get.subcontracts(*args))

    except Exception as e:
        print(e)

    finally:
        return(materials,operations,subcontracts)
