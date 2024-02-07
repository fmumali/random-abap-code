  METHOD set_pane_top_left_cell.
    DATA lv_column_int TYPE zexcel_cell_column.
    DATA lv_row TYPE zexcel_cell_row.

    " Validate input value
    zcl_excel_common=>convert_columnrow2column_a_row(
      EXPORTING
        i_columnrow  = iv_columnrow
      IMPORTING
        e_column_int = lv_column_int
        e_row        = lv_row ).
    IF lv_column_int NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_col AND zcl_excel_common=>c_excel_sheet_max_col
        OR lv_row NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_row AND zcl_excel_common=>c_excel_sheet_max_row.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Invalid column/row coordinates (valid values: A1 to XFD1048576)'.
    ENDIF.
    pane_top_left_cell = iv_columnrow.
  ENDMETHOD.