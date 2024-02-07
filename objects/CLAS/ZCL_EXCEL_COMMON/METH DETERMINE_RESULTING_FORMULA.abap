  METHOD determine_resulting_formula.

    DATA: lv_row_difference TYPE i,
          lv_col_difference TYPE i.

*--------------------------------------------------------------------*
* Calculate distance of reference and current cell
*--------------------------------------------------------------------*
    calculate_cell_distance( EXPORTING
                               iv_reference_cell = iv_reference_cell
                               iv_current_cell   = iv_current_cell
                             IMPORTING
                               ev_row_difference = lv_row_difference
                               ev_col_difference = lv_col_difference ).

*--------------------------------------------------------------------*
* and shift formula by using the row- and columndistance
*--------------------------------------------------------------------*
    ev_resulting_formula = shift_formula( iv_reference_formula = iv_reference_formula
                                          iv_shift_rows        = lv_row_difference
                                          iv_shift_cols        = lv_col_difference ).

  ENDMETHOD.                    "determine_resulting_formula