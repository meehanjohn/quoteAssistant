import pandas as pd

def carry_over(materials,operations,subcontracts,output_file):
    ref_file = r'C:\Users\Amuneal\Work Programs\quoteAssistant\docs\operation_reference.csv'
    ref_df = pd.read_csv(ref_file)
    convert_dict = pd.Series(ref_df.epicor_op_id.values,
                             index=ref_df.vantage_op_desc).to_dict()

    ops_table = pd.DataFrame(operations).dropna(axis=1)
    matls_table = pd.DataFrame(materials).dropna(axis=1)
    subs_table = pd.DataFrame(subcontracts).dropna(axis=1)

    vantage_check = set(['Setup Hours',
                         'Labor Hours',
                         'Laser Hours',
                         'Setup After Sub',
                         'Labor After Sub'])

    ops_set = set(ops_table[0:1].values[0].tolist())
    if vantage_check.issubset(ops_set):
         vantage = True
         ops_table.insert(5,'Empty1',[None]*ops_table.shape[0])
         ops_table.insert(9,'Empty2',[None]*ops_table.shape[0])

    if ops_table.shape[1]>0:
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
    else:
        ops_header=None

    if matls_table.shape[1]>0:
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
    else:
        matls_header=None

    if subs_table.shape[1]>0:
        subs_header = ['col'+str(i) for i in range(subs_table.shape[1]-1)]
        subs_header[:16] = ['subassembly',
                            'row',
                            'op_desc',
                            'resource_grp',
                            'vendor',
                            'col1',
                            'qty',
                            'unit_cost',
                            'ext_cost',
                            'col2',
                            'col3',
                            'col4',
                            'col5',
                            'col6',
                            'col7',
                            'op_id',
                            'resource_id']
    else:
        subs_header=None
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
