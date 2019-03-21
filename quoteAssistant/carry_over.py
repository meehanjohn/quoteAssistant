import pandas as pd

def carry_over(materials,operations,subcontracts,output_file):
    ops_table = pd.DataFrame(operations)
    matls_table = pd.DataFrame(materials)
    subs_table = pd.DataFrame(subcontracts)

    with pd.ExcelWriter(output_file) as writer:
        ops_table.to_excel(writer,sheet_name='operation_entry',header=False,index=False)
        matls_table.to_excel(writer,sheet_name='material_entry',header=False,index=False)
        subs_table.to_excel(writer,sheet_name='subcontract_entry',header=False,index=False)
