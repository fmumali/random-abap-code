  METHOD get_filter_reference.
    DATA: l_row_start_c TYPE string,
          l_row_end_c   TYPE string,
          l_col_start_c TYPE string,
          l_col_end_c   TYPE string,
          l_value       TYPE string.

    validate_area( ).

    l_row_end_c = filter_area-row_end.
    CONDENSE l_row_end_c NO-GAPS.

    l_row_start_c = filter_area-row_start.
    CONDENSE l_row_start_c NO-GAPS.

    l_col_start_c = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_start ) .
    l_col_end_c   = zcl_excel_common=>convert_column2alpha( ip_column = filter_area-col_end ) .
    l_value = worksheet->get_title( ) .

    r_ref = zcl_excel_common=>escape_string( ip_value = l_value ).

    CONCATENATE r_ref '!$' l_col_start_c '$' l_row_start_c ':$' l_col_end_c '$' l_row_end_c INTO r_ref.

  ENDMETHOD.