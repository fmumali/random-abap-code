  METHOD add.

    DATA: ls_autofilter LIKE LINE OF me->mt_autofilters.

    FIELD-SYMBOLS: <ls_autofilter> LIKE LINE OF me->mt_autofilters.

    READ TABLE me->mt_autofilters ASSIGNING <ls_autofilter> WITH TABLE KEY worksheet = io_sheet.
    IF sy-subrc = 0.
      RAISE EXCEPTION TYPE zcx_excel. " adding another autofilter to sheet is not allowed
    ENDIF.

    CREATE OBJECT ro_autofilter
      EXPORTING
        io_sheet = io_sheet.

    ls_autofilter-worksheet  = io_sheet.
    ls_autofilter-autofilter = ro_autofilter.
    INSERT ls_autofilter INTO TABLE me->mt_autofilters.


  ENDMETHOD.