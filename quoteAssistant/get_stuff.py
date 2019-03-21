import win32com.client as win32

def materials(sheet_index,sheet_name,sheet_obj):
    range = sheet_obj.UsedRange.Rows
    sheet_values = [list(row.Value2[0]) for row in range if type(row.Value2) is tuple]
    values_dict = dict(enumerate(sheet_values))

    new_matlstart_key = set(['Extended Cost'])
    old_matlstart_key = set(['Description',
                             'Req\'d to Start Laser or Fab (F, L)',
                             'Vendor',
                             'Days        Lead-Time',
                             'Qty Req\'d'])

    matlend_key = set(['Sub-Contract ',
                       'Days'])

    try:
        mtl_st_indcs = [k for k,v in values_dict.items()
                        if old_matlstart_key.issubset(set(v))
                        or (new_matlstart_key.issubset(set(v))
                        and 'Description' in v[0])]

        mtl_end_indcs = [k for k,v in values_dict.items()
                         if matlend_key.issubset(set(v))]

        materials = [[sheet_name]+[k]+v for k,v in values_dict.items()
                      if k > mtl_st_indcs[0]
                      and k < mtl_end_indcs[0]
                      and isinstance(v[0],str)]

    except:
        materials = []

    finally:
        return materials

def operations(sheet_index,sheet_name,sheet_obj):
    range = sheet_obj.UsedRange.Rows
    sheet_values = [list(row.Value2[0]) for row in range if type(row.Value2) is tuple]
    values_dict = dict(enumerate(sheet_values))

    new_opstart_key = set(['Epicore Operations',
                       'Epicore Resource Group',
                       'Qty',
                       'Set-up',
                       'Standard'])

    old_opstart_key = set(['Operation',
                       'B / A',
                       'Qty',
                       'Set-up',
                       'Rate'])

    opend_key = set(['Total Labor / Set-up Costs'])

    try:
        op_st_indcs = [k for k,v in values_dict.items()
                       if new_opstart_key.issubset(set(v))
                       or old_opstart_key.issubset(set(v))]

        op_end_indcs = [k for k,v in values_dict.items()
                        if opend_key.issubset(set(v))]

        operations = [[sheet_name]+[k]+v for k,v in values_dict.items()
                  if k > op_st_indcs[0]
                  and k < op_end_indcs[0]
                  and isinstance(v[2],float)
                  and v[2] > 0]
    except:
        operations = []

    finally:
        return operations

def subcontracts(sheet_index,sheet_name,sheet_obj):
    range = sheet_obj.UsedRange.Rows
    sheet_values = [list(row.Value2[0]) for row in range if type(row.Value2) is tuple]
    values_dict = dict(enumerate(sheet_values))

    new_substart_key = set(['Description',
                            'Epicore Resource Group',
                            'Vendor Name',
                            'Lead-Time',
                            'Qty',
                            'Unit Cost'])

    old_substart_key = set(['Description',
                            None,
                            'Vendor',
                            'Lead-Time',
                            'Qty',
                            'Unit Cost'])

    subend_key = set(['Assembly 1', 'Hours', 'Mins'])

    try:
        sub_st_indcs = [k for k,v in values_dict.items()
                        if new_substart_key.issubset(set(v))
                        or old_substart_key.issubset(set(v))]

        sub_end_indcs = [k-4 for k,v in values_dict.items()
                         if subend_key.issubset(set(v))]

        subcontracts = [[sheet_name]+[k]+v for k,v in values_dict.items()
                         if k > sub_st_indcs[0]
                         and k < sub_end_indcs[0]
                         and isinstance(v[4],float)]

    except:
        subcontracts = []

    finally:
        return subcontracts
