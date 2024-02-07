  METHOD get_active_worksheet.

    eo_worksheet = me->worksheets->get( me->worksheets->active_worksheet ).

  ENDMETHOD.