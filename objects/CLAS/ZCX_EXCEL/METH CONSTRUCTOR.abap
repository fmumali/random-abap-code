  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    CALL METHOD super->constructor
      EXPORTING
        textid   = textid
        previous = previous.
    IF textid IS INITIAL.
      me->textid = zcx_excel .
    ENDIF.
    me->error = error .
    me->syst_at_raise = syst_at_raise .
  ENDMETHOD.