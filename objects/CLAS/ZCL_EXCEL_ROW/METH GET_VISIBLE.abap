  METHOD get_visible.

    DATA: lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
    FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF lt_row_outlines.

    r_visible = me->visible.
    CHECK r_visible = abap_true.  " Currently visible --> but maybe the new outline methodology will hide it implicitly
    CHECK io_worksheet IS BOUND.  " But we have to see the worksheet to make sure

    lt_row_outlines = io_worksheet->get_row_outlines( ).
    LOOP AT lt_row_outlines ASSIGNING <ls_row_outline> WHERE row_from  <= me->row_index
                                                         AND row_to    >= me->row_index
                                                         AND collapsed =  abap_true.      " row is in a collapsed outline --> not visible
      CLEAR r_visible.
      RETURN. " one hit is enough to ensure invisibility

    ENDLOOP.

  ENDMETHOD.