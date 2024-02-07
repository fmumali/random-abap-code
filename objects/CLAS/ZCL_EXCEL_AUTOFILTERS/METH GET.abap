  METHOD get.

    FIELD-SYMBOLS: <ls_autofilter> LIKE LINE OF me->mt_autofilters.

    READ TABLE me->mt_autofilters ASSIGNING <ls_autofilter> WITH TABLE KEY worksheet = io_worksheet.
    IF sy-subrc = 0.
      ro_autofilter = <ls_autofilter>-autofilter.
    ELSE.
      CLEAR ro_autofilter.
    ENDIF.

  ENDMETHOD.