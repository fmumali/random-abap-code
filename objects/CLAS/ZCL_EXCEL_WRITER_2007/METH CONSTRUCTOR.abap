  METHOD constructor.
    DATA: lt_unicode_point_codes TYPE TABLE OF string,
          lv_unicode_point_code  TYPE i.

    me->ixml = cl_ixml=>create( ).

    SPLIT '0,1,2,3,4,5,6,7,8,' " U+0000 to U+0008
       && '11,12,'             " U+000B, U+000C
       && '14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,' " U+000E to U+001F
       && '65534,65535'        " U+FFFE, U+FFFF
      AT ',' INTO TABLE lt_unicode_point_codes.
    control_characters = ``.
    LOOP AT lt_unicode_point_codes INTO lv_unicode_point_code.
      control_characters = control_characters && cl_abap_conv_in_ce=>uccpi( lv_unicode_point_code ).
    ENDLOOP.

  ENDMETHOD.