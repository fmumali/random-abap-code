  METHOD update_dimension_range.

    DATA: ls_sheet_content TYPE zexcel_s_cell_data,
          lv_row_alpha     TYPE string,
          lv_column_alpha  TYPE zexcel_cell_column_alpha.

    CHECK sheet_content IS NOT INITIAL.

    upper_cell-cell_row = rows->get_min_index( ).
    IF upper_cell-cell_row = 0.
      upper_cell-cell_row = zcl_excel_common=>c_excel_sheet_max_row.
    ENDIF.
    upper_cell-cell_column = zcl_excel_common=>c_excel_sheet_max_col.

    lower_cell-cell_row = rows->get_max_index( ).
    IF lower_cell-cell_row = 0.
      lower_cell-cell_row = zcl_excel_common=>c_excel_sheet_min_row.
    ENDIF.
    lower_cell-cell_column = zcl_excel_common=>c_excel_sheet_min_col.

    LOOP AT sheet_content INTO ls_sheet_content.
      IF upper_cell-cell_row > ls_sheet_content-cell_row.
        upper_cell-cell_row = ls_sheet_content-cell_row.
      ENDIF.
      IF upper_cell-cell_column > ls_sheet_content-cell_column.
        upper_cell-cell_column = ls_sheet_content-cell_column.
      ENDIF.
      IF lower_cell-cell_row < ls_sheet_content-cell_row.
        lower_cell-cell_row = ls_sheet_content-cell_row.
      ENDIF.
      IF lower_cell-cell_column < ls_sheet_content-cell_column.
        lower_cell-cell_column = ls_sheet_content-cell_column.
      ENDIF.
    ENDLOOP.

    upper_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = upper_cell-cell_column i_row = upper_cell-cell_row ).

    lower_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lower_cell-cell_column i_row = lower_cell-cell_row ).

  ENDMETHOD.                    "UPDATE_DIMENSION_RANGE