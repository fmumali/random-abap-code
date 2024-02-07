  METHOD get_right_column_integer.
    DATA: ls_field_catalog  TYPE zexcel_s_fieldcatalog.

    IF settings-bottom_right_column IS NOT INITIAL.
      ev_column =  zcl_excel_common=>convert_column2int( settings-bottom_right_column ).
      RETURN.
    ENDIF.

    ev_column =  zcl_excel_common=>convert_column2int( settings-top_left_column ).
    LOOP AT fieldcat INTO ls_field_catalog WHERE dynpfld EQ abap_true.
      ADD 1 TO ev_column.
    ENDLOOP.

  ENDMETHOD.