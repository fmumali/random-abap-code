  METHOD is_empty.
    IF me->mt_autofilters IS INITIAL.
      r_empty = abap_true.
    ENDIF.
  ENDMETHOD.