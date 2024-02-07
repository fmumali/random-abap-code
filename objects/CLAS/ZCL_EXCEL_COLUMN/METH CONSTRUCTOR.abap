  METHOD constructor.
    me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
    me->width         = -1.
    me->auto_size     = abap_false.
    me->visible       = abap_true.
    me->outline_level = 0.
    me->collapsed     = abap_false.
    me->excel         = ip_excel.        "ins issue #157 - Allow Style for columns
    me->worksheet     = ip_worksheet.    "ins issue #157 - Allow Style for columns

    " set default index to cellXf
    me->xf_index = 0.

  ENDMETHOD.