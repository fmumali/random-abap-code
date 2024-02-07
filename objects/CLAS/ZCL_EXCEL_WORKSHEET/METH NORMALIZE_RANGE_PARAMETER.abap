  METHOD normalize_range_parameter.

    DATA: lv_errormessage TYPE string.

    IF ( ( ip_column_start IS NOT INITIAL OR ip_column_end IS NOT INITIAL
            OR ip_row IS NOT INITIAL OR ip_row_to IS NOT INITIAL ) AND ip_range IS NOT INITIAL )
        OR ( ip_column_start IS INITIAL AND ip_column_end IS INITIAL
            AND ip_row IS INITIAL AND ip_row_to IS INITIAL AND ip_range IS INITIAL ).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Please provide either row and column interval, or range reference'.
    ENDIF.

    IF ip_range IS NOT INITIAL.
      zcl_excel_common=>convert_range2column_a_row(
        EXPORTING
          i_range            = ip_range
        IMPORTING
          e_column_start_int = ep_column_start
          e_column_end_int   = ep_column_end
          e_row_start        = ep_row
          e_row_end          = ep_row_to ).
    ELSE.
      IF ip_column_start IS INITIAL.
        ep_column_start = zcl_excel_common=>c_excel_sheet_min_col.
      ELSE.
        ep_column_start = zcl_excel_common=>convert_column2int( ip_column_start ).
      ENDIF.
      IF ip_column_end IS INITIAL.
        ep_column_end = ep_column_start.
      ELSE.
        ep_column_end = zcl_excel_common=>convert_column2int( ip_column_end ).
      ENDIF.
      ep_row = ip_row.
      IF ep_row IS INITIAL.
        ep_row = zcl_excel_common=>c_excel_sheet_min_row.
      ENDIF.
      ep_row_to = ip_row_to.
      IF ep_row_to IS INITIAL.
        ep_row_to = ep_row.
      ENDIF.
    ENDIF.

    IF ep_row > ep_row_to.
      lv_errormessage = 'First row larger than last row'(405).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF ep_column_start > ep_column_end.
      lv_errormessage = 'First column larger than last column'(406).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

  ENDMETHOD.