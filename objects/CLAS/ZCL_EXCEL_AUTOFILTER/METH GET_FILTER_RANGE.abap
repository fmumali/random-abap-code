  METHOD get_filter_range.
    DATA: l_row_start_c TYPE string,
          l_row_end_c   TYPE string,
          l_col_start_c TYPE string,
          l_col_end_c   TYPE string.

    validate_area( ).

    l_row_end_c = filter_area-row_end.
    CONDENSE l_row_end_c NO-GAPS.

    l_row_start_c = filter_area-row_start.
    CONDENSE l_row_start_c NO-GAPS.

    l_col_start_c = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_start ) .
    l_col_end_c   = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_end ) .

    CONCATENATE l_col_start_c l_row_start_c ':' l_col_end_c l_row_end_c INTO r_range.

  ENDMETHOD.