  METHOD get_reference.
    DATA: lv_column                TYPE zexcel_cell_column,
          lv_table_lines           TYPE i,
          lv_right_column          TYPE zexcel_cell_column_alpha,
          ls_field_catalog         TYPE zexcel_s_fieldcatalog,
          lv_bottom_row            TYPE zexcel_cell_row,
          lv_top_row_string(10)    TYPE c,
          lv_bottom_row_string(10) TYPE c.

    FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE.

*column
    lv_column = zcl_excel_common=>convert_column2int( settings-top_left_column ).
    lv_table_lines = 0.
    LOOP AT fieldcat INTO ls_field_catalog WHERE dynpfld EQ abap_true.
      ADD 1 TO lv_table_lines.
    ENDLOOP.
    lv_column = lv_column + lv_table_lines - 1.
    lv_right_column  = zcl_excel_common=>convert_column2alpha( lv_column ).

*row
    ASSIGN table_data->* TO <fs_table>.
    DESCRIBE TABLE <fs_table> LINES lv_table_lines.
    IF lv_table_lines = 0.
      lv_table_lines = 1. "table needs at least 1 data row
    ENDIF.
    lv_bottom_row = settings-top_left_row + lv_table_lines .

    IF me->has_totals( ) = abap_true AND ip_include_totals_row = abap_true.
      ADD 1 TO lv_bottom_row.
    ENDIF.

    lv_top_row_string = zcl_excel_common=>number_to_excel_string( settings-top_left_row ).
    lv_bottom_row_string = zcl_excel_common=>number_to_excel_string( lv_bottom_row ).

    CONCATENATE settings-top_left_column lv_top_row_string
                ':'
                lv_right_column lv_bottom_row_string INTO ov_reference.

  ENDMETHOD.