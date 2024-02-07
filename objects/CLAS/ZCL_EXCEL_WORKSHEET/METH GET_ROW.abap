  METHOD get_row.
    eo_row = me->rows->get( ip_index = ip_row ).

    IF eo_row IS NOT BOUND.
      eo_row = me->add_new_row( ip_row ).
    ENDIF.
  ENDMETHOD.                    "GET_ROW