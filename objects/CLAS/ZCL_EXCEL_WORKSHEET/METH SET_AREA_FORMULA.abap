  METHOD set_area_formula.
    DATA: ld_row              TYPE zexcel_cell_row,
          ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column           TYPE zexcel_cell_column_alpha,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    IF ip_merge = abap_true OR ip_area = c_area-topleft.

      me->set_cell_formula( ip_column = ld_column_start_int ip_row = ld_row_start
                            ip_formula = ip_formula ).

    ELSE.

      ld_column_int = ld_column_start_int.
      WHILE ld_column_int <= ld_column_end_int.

        ld_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
        ld_row = ld_row_start.
        WHILE ld_row <= ld_row_end.

          me->set_cell_formula( ip_column = ld_column ip_row = ld_row
                                ip_formula = ip_formula ).

          ADD 1 TO ld_row.
        ENDWHILE.

        ADD 1 TO ld_column_int.
      ENDWHILE.

    ENDIF.

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "set_area_formula