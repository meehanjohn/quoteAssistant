import pandas as pd
from quoteAssistant import *
import win32com.client as win32
from gooey import Gooey, GooeyParser
import os.path
import time

def main():
    """ Inputs are currently defined by maintaining a simple CSV
    containing the Sales Order number and server filepath where
    quotes are located."""

    path = os.path.abspath(os.path.dirname(__file__))

    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False
        input_csv = excel.Workbooks.Open(str(path)+'\\..\\docs\\job_entry.csv')
        excel.Visible = True

    except Exception as e:
        print(e)
        quit()

    excel_open = bool(excel.Workbooks.Count)

    while excel_open:
        try:
            excel_open = bool(excel.Workbooks.Count)
        except:
            time.sleep(5)

    df = pd.read_csv(str(path)+'\\..\\docs\\job_entry.csv')

    for index, row in df.iterrows():
        job_files = create_job_list(row[0], row[1])

        for input_file, output_file in job_files:

            quote_book = load_spreadsheet(input_file)
            materials, operations, subcontracts = read_file(quote_book)

            job_entry_book = load_spreadsheet()
            job_entry_book.SaveAs(output_file)
            job_entry_book.Close(True)

            carry_over(materials,operations,subcontracts,output_file)
            make_table(output_file)

            print("{} complete".format(output_file))

if __name__ == "__main__":
	main()
