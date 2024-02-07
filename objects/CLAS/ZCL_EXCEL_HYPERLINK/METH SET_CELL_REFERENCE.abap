  METHOD set_cell_reference.
    me->column = zcl_excel_common=>convert_column2alpha( ip_column ). " issue #155 - less restrictive typing for ip_column
    me->row = ip_row.
  ENDMETHOD.