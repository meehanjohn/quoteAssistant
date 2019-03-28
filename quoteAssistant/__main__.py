import pandas as pd
from quoteAssistant import *
import win32com.client as win32
from gooey import Gooey, GooeyParser

#@Gooey
def main():
    """ Inputs are currently defined by maintaining a simple CSV
    containing the Sales Order number and server filepath where
    quotes are located."""

    #parser = GooeyParser(description="Quote Assistant")
    #parser.add_argument('Input_Directory',
                        #help="Location of Quote(s):",
                        #widget='DirChooser')
    #parser.add_argument('SO_Number',
                        #widget='TextField')

    #args = parser.parse_args()


    df = pd.read_csv('docs\\job_entry.csv')

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
