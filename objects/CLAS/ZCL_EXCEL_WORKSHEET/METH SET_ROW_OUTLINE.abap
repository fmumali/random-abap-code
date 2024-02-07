  METHOD set_row_outline.

    DATA: ls_row_outline LIKE LINE OF me->mt_row_outlines.
    FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF me->mt_row_outlines.

    READ TABLE me->mt_row_outlines ASSIGNING <ls_row_outline> WITH TABLE KEY row_from = iv_row_from
                                                                             row_to   = iv_row_to.
    IF sy-subrc <> 0.
      IF iv_row_from <= 0.
        zcx_excel=>raise_text( 'First row of outline must be a positive number' ).
      ENDIF.
      IF iv_row_to < iv_row_from.
        zcx_excel=>raise_text( 'Last row of outline may not be less than first line of outline' ).
      ENDIF.
      ls_row_outline-row_from = iv_row_from.
      ls_row_outline-row_to   = iv_row_to.
      INSERT ls_row_outline INTO TABLE me->mt_row_outlines ASSIGNING <ls_row_outline>.
    ENDIF.

    CASE iv_collapsed.

      WHEN abap_true
        OR abap_false.
        <ls_row_outline>-collapsed = iv_collapsed.

      WHEN OTHERS.
        zcx_excel=>raise_text( 'Unknown collapse state' ).

    ENDCASE.
  ENDMETHOD.                    "SET_ROW_OUTLINE