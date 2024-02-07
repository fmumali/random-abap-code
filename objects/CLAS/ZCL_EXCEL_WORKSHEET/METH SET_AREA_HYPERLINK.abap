  METHOD set_area_hyperlink.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row,
          ld_value            TYPE string,
          ld_formula          TYPE string.
    DATA: lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    ld_column_int = ld_column_start_int.
    WHILE ld_column_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.

        me->get_cell( EXPORTING ip_column  = ld_current_column ip_row = ld_current_row
                      IMPORTING ep_value   = ld_value
                                ep_formula = ld_formula ).

        IF ip_is_internal = abap_true.
          lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = ip_url ).
        ELSE.
          lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = ip_url ).
        ENDIF.

        me->set_cell( ip_column = ld_current_column ip_row = ld_current_row ip_value = ld_value ip_formula = ld_formula ip_hyperlink = lo_hyperlink ).

        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_int.
    ENDWHILE.

  ENDMETHOD.                    "SET_AREA_HYPERLINK