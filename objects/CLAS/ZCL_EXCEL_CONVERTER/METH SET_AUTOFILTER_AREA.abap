  METHOD set_autofilter_area.
    DATA: ls_area   TYPE zexcel_s_autofilter_area,
          l_lines   TYPE i,
          lt_values TYPE zexcel_t_autofilter_values,
          ls_values TYPE zexcel_s_autofilter_values.

* Let's check for filter.
    IF wo_autofilter IS BOUND.
      ls_area-row_start = 1.
      lt_values = wo_autofilter->get_values( ) .
      SORT lt_values BY column ASCENDING.
      DESCRIBE TABLE lt_values LINES l_lines.
      READ TABLE lt_values INTO ls_values INDEX 1.
      IF sy-subrc = 0.
        ls_area-col_start = ls_values-column.
      ENDIF.
      READ TABLE lt_values INTO ls_values INDEX l_lines.
      IF sy-subrc = 0.
        ls_area-col_end = ls_values-column.
      ENDIF.
      wo_autofilter->set_filter_area( is_area = ls_area ) .
    ENDIF.

  ENDMETHOD.