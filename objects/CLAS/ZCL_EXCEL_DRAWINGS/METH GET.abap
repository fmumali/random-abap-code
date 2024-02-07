  METHOD get.

    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_drawing ?= drawings->get( lv_index ).
  ENDMETHOD.