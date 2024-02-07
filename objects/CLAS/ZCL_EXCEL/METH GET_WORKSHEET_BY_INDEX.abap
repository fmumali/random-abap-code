  METHOD get_worksheet_by_index.


    DATA: lv_index TYPE zexcel_active_worksheet.

    lv_index = iv_index.
    eo_worksheet = me->worksheets->get( lv_index ).

  ENDMETHOD.