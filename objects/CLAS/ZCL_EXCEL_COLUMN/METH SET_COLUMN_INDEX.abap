  METHOD set_column_index.
    me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
    io_column = me.
  ENDMETHOD.