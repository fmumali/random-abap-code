  METHOD is_cell_in_range.
    DATA lv_column_start    TYPE zexcel_cell_column_alpha.
    DATA lv_column_end      TYPE zexcel_cell_column_alpha.
    DATA lv_row_start       TYPE zexcel_cell_row.
    DATA lv_row_end         TYPE zexcel_cell_row.
    DATA lv_column_start_i  TYPE zexcel_cell_column.
    DATA lv_column_end_i    TYPE zexcel_cell_column.
    DATA lv_column_i        TYPE zexcel_cell_column.


* Split range and convert columns
    convert_range2column_a_row(
      EXPORTING
        i_range        = ip_range
      IMPORTING
        e_column_start = lv_column_start
        e_column_end   = lv_column_end
        e_row_start    = lv_row_start
        e_row_end      = lv_row_end ).

    lv_column_start_i = convert_column2int( ip_column = lv_column_start ).
    lv_column_end_i   = convert_column2int( ip_column = lv_column_end ).

    lv_column_i = convert_column2int( ip_column = ip_column ).

* Check if cell is in range
    IF lv_column_i >= lv_column_start_i AND
       lv_column_i <= lv_column_end_i   AND
       ip_row      >= lv_row_start      AND
       ip_row      <= lv_row_end.
      rp_in_range = abap_true.
    ENDIF.
  ENDMETHOD.