  METHOD get_title.
    DATA lv_value TYPE string.
    IF ip_escaped EQ abap_true.
      lv_value = me->title.
      ep_title = zcl_excel_common=>escape_string( lv_value ).
    ELSE.
      ep_title = me->title.
    ENDIF.
  ENDMETHOD.                    "GET_TITLE