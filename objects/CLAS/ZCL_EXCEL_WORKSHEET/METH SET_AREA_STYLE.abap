  METHOD set_area_style.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row.

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
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_int.
    ENDWHILE.
    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "SET_AREA_STYLE