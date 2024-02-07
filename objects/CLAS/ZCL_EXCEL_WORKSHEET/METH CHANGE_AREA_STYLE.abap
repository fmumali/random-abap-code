  METHOD change_area_style.

    DATA: lv_row              TYPE zexcel_cell_row,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_to           TYPE zexcel_cell_row,
          lv_column_int       TYPE zexcel_cell_column,
          lv_column_start_int TYPE zexcel_cell_column,
          lv_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_to ).

    lv_column_int = lv_column_start_int.
    WHILE lv_column_int <= lv_column_end_int.

      lv_row = lv_row_start.
      WHILE lv_row <= lv_row_to.

        ip_style_changer->apply( ip_worksheet = me
                                 ip_column    = lv_column_int
                                 ip_row       = lv_row ).

        ADD 1 TO lv_row.
      ENDWHILE.

      ADD 1 TO lv_column_int.
    ENDWHILE.

  ENDMETHOD.