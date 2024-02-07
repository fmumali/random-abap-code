  METHOD calculate_table_bottom_right.

    DATA: lv_errormessage TYPE string,
          lv_columns      TYPE i,
          lt_columns      TYPE zexcel_t_fieldcatalog,
          lv_maxrow       TYPE i,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_curtable     TYPE REF TO zcl_excel_table,
          lv_row_int      TYPE zexcel_cell_row,
          lv_column_int   TYPE zexcel_cell_column,
          lv_rows         TYPE i,
          lv_maxcol       TYPE i.

    "Get the number of columns for the current table
    lt_columns = it_field_catalog.
    DELETE lt_columns WHERE dynpfld NE abap_true.
    DESCRIBE TABLE lt_columns LINES lv_columns.

    "Calculate the top left row of the current table
    lv_column_int = zcl_excel_common=>convert_column2int( cs_settings-top_left_column ).
    lv_row_int    = cs_settings-top_left_row.

    "Get number of row for the current table
    DESCRIBE TABLE ip_table LINES lv_rows.

    "Calculate the bottom right row for the current table
    lv_maxcol                       = lv_column_int + lv_columns - 1.
    lv_maxrow                       = lv_row_int    + lv_rows.
    cs_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( lv_maxcol ).
    cs_settings-bottom_right_row    = lv_maxrow.

  ENDMETHOD.