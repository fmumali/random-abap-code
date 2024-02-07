  METHOD get_color.
    DATA: lv_index TYPE i.

    lv_index = ip_index + 1.
    READ TABLE colors INTO ep_color INDEX lv_index.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'Invalid color index' ).
    ENDIF.
  ENDMETHOD.