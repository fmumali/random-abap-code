  METHOD set_position.
    from_loc-col = zcl_excel_common=>convert_column2int( ip_from_col ) - 1.
    IF ip_coloff IS SUPPLIED.
      from_loc-col_offset = ip_coloff.
    ENDIF.
    from_loc-row = ip_from_row - 1.
    IF ip_rowoff IS SUPPLIED.
      from_loc-row_offset = ip_rowoff.
    ENDIF.
    anchor = anchor_one_cell.
  ENDMETHOD.