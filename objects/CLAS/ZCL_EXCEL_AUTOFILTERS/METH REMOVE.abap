  METHOD remove.

    DATA: lo_worksheet  TYPE REF TO zcl_excel_worksheet.

    DELETE TABLE me->mt_autofilters WITH TABLE KEY worksheet = lo_worksheet.

  ENDMETHOD.