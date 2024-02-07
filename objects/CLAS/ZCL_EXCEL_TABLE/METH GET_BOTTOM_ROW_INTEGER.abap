  METHOD get_bottom_row_integer.
    DATA: lv_table_lines TYPE i.
    FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE.

    IF settings-bottom_right_row IS NOT INITIAL.
*    ev_row =  zcl_excel_common=>convert_column2int( settings-bottom_right_row ). " del issue #246
      ev_row =  settings-bottom_right_row .                                         " ins issue #246
      RETURN.
    ENDIF.

    ASSIGN table_data->* TO <fs_table>.
    DESCRIBE TABLE <fs_table> LINES lv_table_lines.
    IF lv_table_lines = 0.
      lv_table_lines = 1. "table needs at least 1 data row
    ENDIF.

    ev_row = settings-top_left_row + lv_table_lines.

    IF me->has_totals( ) = abap_true."  ????  AND ip_include_totals_row = abap_true.
      ADD 1 TO ev_row.
    ENDIF.
  ENDMETHOD.