  METHOD get.
    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_style_cond ?= styles_cond->get( lv_index ).
  ENDMETHOD.