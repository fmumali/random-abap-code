  METHOD get.

    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_worksheet ?= worksheets->get( lv_index ).

  ENDMETHOD.