import pandas as pd

def carry_over(materials,operations,subcontracts,output_file):
    ops_table = pd.DataFrame(operations)
    matls_table = pd.DataFrame(materials)
    subs_table = pd.DataFrame(subcontracts)

    ops_header = ['col'+str(i) for i in range(ops_table.shape[1]-1)]
    ops_header[:18] = ['subassembly',
                       'row',
                       'op_desc',
                       'resource_grp',
                       'qty',
                       'col1',
                       'col2',
                       'col3',
                       'col4',
                       'col5',
                       'col6',
                       'col7',
                       'col8',
                       'col9',
                       'op_id',
                       'col10',
                       'setup',
                       'labor',
                       'col11']

    matls_header = ['col'+str(i) for i in range(matls_table.shape[1]-1)]
    matls_header[:11] = ['subassembly',
                         'row',
                         'desc1',
                         'desc2',
                         'col1',
                         'col2',
                         'qty',
                         'col4',
                         'col5',
                         'col6',
                         'ext_cost',
                         'comment']

    subs_header = ['col'+str(i) for i in range(subs_table.shape[1]-1)]
    subs_header = None
    #TODO: figure out proper subcontract header tables


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
