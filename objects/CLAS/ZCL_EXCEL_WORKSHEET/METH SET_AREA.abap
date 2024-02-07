  METHOD set_area.

    DATA: lv_row              TYPE zexcel_cell_row,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_end          TYPE zexcel_cell_row,
          lv_column_int       TYPE zexcel_cell_column,
          lv_column           TYPE zexcel_cell_column_alpha,
          lv_column_start_int TYPE zexcel_cell_column,
          lv_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    IF ip_merge = abap_true OR ip_area = c_area-topleft.

      IF ip_data_type IS SUPPLIED OR
         ip_abap_type IS SUPPLIED.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink
                      ip_data_type = ip_data_type
                      ip_abap_type = ip_abap_type ).

      ELSE.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink ).

      ENDIF.

    ELSE.

      lv_column_int = lv_column_start_int.
      WHILE lv_column_int <= lv_column_end_int.

        lv_column = zcl_excel_common=>convert_column2alpha( lv_column_int ).
        lv_row = lv_row_start.

        WHILE lv_row <= lv_row_end.

          IF ip_data_type IS SUPPLIED OR
             ip_abap_type IS SUPPLIED.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink
                          ip_data_type = ip_data_type
                          ip_abap_type = ip_abap_type ).

          ELSE.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink ).

          ENDIF.

          ADD 1 TO lv_row.
        ENDWHILE.

        ADD 1 TO lv_column_int.
      ENDWHILE.

    ENDIF.

    IF ip_style IS SUPPLIED.

      me->set_area_style( ip_column_start = lv_column_start_int
                          ip_column_end   = lv_column_end_int
                          ip_row          = lv_row_start
                          ip_row_to       = lv_row_end
                          ip_style        = ip_style ).
    ENDIF.

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.

      me->set_merge( ip_column_start = lv_column_start_int
                     ip_column_end   = lv_column_end_int
                     ip_row          = lv_row_start
                     ip_row_to       = lv_row_end ).

    ENDIF.

  ENDMETHOD.                    "set_area