import pandas as pd

def carry_over(materials,operations,subcontracts,output_file):
    ops_table = pd.DataFrame(operations)
    matls_table = pd.DataFrame(materials)
    subs_table = pd.DataFrame(subcontracts)

    ops_header = ['Subassembly',
                  'Row',
                  'Op Desc',
                  'Resource Group',
                  'Qty',
                  'Col1',
                  'Col2',
                  'Col3',
                  'Col4',
                  'Col5',
                  'Col6',
                  'Col7',
                  'Col8',
                  'Col9',
                  'Op ID',
                  'Col10'
                  'Setup'
                  'Labor',
                  'Col11']
                  
    matls_header = ['Subassembly',
                    'Row',
                    'Desc1',
                    'Desc2',
                    'Col1',
                    'Col2',
                    'Qty',
                    'Col4',
                    'Col5',
                    'Col6',
                    'Ext Cost',
                    'Comment']
    subs_header = None


    with pd.ExcelWriter(output_file) as writer:
        ops_table.to_excel(writer,
                           sheet_name='operation_entry',
                           header=ops_header,
                           index=False)
        matls_table.to_excel(writer,
                             sheet_name='material_entry',
                             header=matls_header,
                             index=False)
        subs_table.to_excel(writer,
                            sheet_name='subcontract_entry',
                            header=subs_header,
                            index=False)
