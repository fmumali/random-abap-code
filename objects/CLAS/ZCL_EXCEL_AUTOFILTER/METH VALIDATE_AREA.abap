  METHOD validate_area.
    DATA: l_col                   TYPE zexcel_cell_column,
          ls_original_filter_area TYPE zexcel_s_autofilter_area,
          l_row                   TYPE zexcel_cell_row.

    l_row = worksheet->get_highest_row( ) .
    l_col = worksheet->get_highest_column( ) .

    IF filter_area IS INITIAL.
      filter_area-row_start = 1.
      filter_area-col_start = 1.
      filter_area-row_end   = l_row .
      filter_area-col_end   = l_col .
    ENDIF.

    IF filter_area-row_start > filter_area-row_end.
      ls_original_filter_area = filter_area.
      filter_area-row_start = ls_original_filter_area-row_end.
      filter_area-row_end = ls_original_filter_area-row_start.
    ENDIF.
    IF filter_area-row_start < 1.
      filter_area-row_start = 1.
    ENDIF.
    IF filter_area-col_start < 1.
      filter_area-col_start = 1.
    ENDIF.
    IF filter_area-row_end > l_row OR
       filter_area-row_end < 1.
      filter_area-row_end = l_row.
    ENDIF.
    IF filter_area-col_end > l_col OR
       filter_area-col_end < 1.
      filter_area-col_end = l_col.
    ENDIF.
    IF filter_area-col_start > filter_area-col_end.
      filter_area-col_start = filter_area-col_end.
    ENDIF.
  ENDMETHOD.