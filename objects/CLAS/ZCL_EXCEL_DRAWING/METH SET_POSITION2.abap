  METHOD set_position2.

    DATA: lv_anchor                     TYPE zexcel_drawing_anchor.
    lv_anchor = ip_anchor.

    IF lv_anchor IS INITIAL.
      IF ip_to IS NOT INITIAL.
        lv_anchor = anchor_two_cell.
      ELSE.
        lv_anchor = anchor_one_cell.
      ENDIF.
    ENDIF.

    CASE lv_anchor.
      WHEN anchor_absolute OR anchor_one_cell.
        CLEAR: me->to_loc.
      WHEN anchor_two_cell.
        CLEAR: me->size.
    ENDCASE.

    me->from_loc = ip_from.
    me->to_loc = ip_to.
    me->anchor = lv_anchor.

  ENDMETHOD.