  METHOD get_outline_level.

    DATA: lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
    FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF lt_row_outlines.

* if someone has set the outline level explicitly - just use that
    IF me->outline_level IS NOT INITIAL.
      r_outline_level = me->outline_level.
      RETURN.
    ENDIF.
* Maybe we can use the outline information in the worksheet
    CHECK io_worksheet IS BOUND.

    lt_row_outlines = io_worksheet->get_row_outlines( ).
    LOOP AT lt_row_outlines ASSIGNING <ls_row_outline> WHERE row_from <= me->row_index
                                                         AND row_to   >= me->row_index.

      ADD 1 TO r_outline_level.

    ENDLOOP.

  ENDMETHOD.