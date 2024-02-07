  METHOD bind_cells.

* Do we need subtotals with grouping
    READ TABLE wt_fieldcatalog TRANSPORTING NO FIELDS WITH KEY is_subtotalled = abap_true.
    IF sy-subrc = 0  .
      r_freeze_col = loop_subtotal( i_row_int = w_row_int
                                    i_col_int  = w_col_int ) .
    ELSE.
      r_freeze_col = loop_normal( i_row_int = w_row_int
                                  i_col_int = w_col_int ) .
    ENDIF.

  ENDMETHOD.