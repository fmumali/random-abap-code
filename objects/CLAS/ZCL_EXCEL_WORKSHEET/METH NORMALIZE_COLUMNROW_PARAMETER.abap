  METHOD normalize_columnrow_parameter.

    IF ( ( ip_column IS NOT INITIAL OR ip_row IS NOT INITIAL ) AND ip_columnrow IS NOT INITIAL )
        OR ( ip_column IS INITIAL AND ip_row IS INITIAL AND ip_columnrow IS INITIAL ).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Please provide either row and column, or cell reference'.
    ENDIF.

    IF ip_columnrow IS NOT INITIAL.
      zcl_excel_common=>convert_columnrow2column_a_row(
        EXPORTING
          i_columnrow  = ip_columnrow
        IMPORTING
          e_column_int = ep_column
          e_row        = ep_row ).
    ELSE.
      ep_column = zcl_excel_common=>convert_column2int( ip_column ).
      ep_row    = ip_row.
    ENDIF.

  ENDMETHOD.