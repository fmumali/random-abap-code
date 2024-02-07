  METHOD set_outline_level.
    IF   ip_outline_level < 0
      OR ip_outline_level > 7.

      zcx_excel=>raise_text( 'Outline level must range between 0 and 7.' ).

    ENDIF.
    me->outline_level = ip_outline_level.
  ENDMETHOD.