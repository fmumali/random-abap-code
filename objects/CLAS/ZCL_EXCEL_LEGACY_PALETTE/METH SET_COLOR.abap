  METHOD set_color.
    DATA: lv_index TYPE i.

    FIELD-SYMBOLS: <lv_color> LIKE LINE OF colors.

    lv_index = ip_index + 1.
    READ TABLE colors ASSIGNING <lv_color> INDEX lv_index.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'Invalid color index' ).
    ENDIF.

    IF <lv_color> <> ip_color.
      modified = abap_true.
      <lv_color> = ip_color.
    ENDIF.

  ENDMETHOD.