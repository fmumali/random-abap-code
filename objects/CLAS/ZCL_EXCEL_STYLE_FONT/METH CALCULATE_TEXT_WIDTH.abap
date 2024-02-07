  METHOD calculate_text_width.
    " Addition to solve issue #120, contribution by Stefan Schmoecker
    r_width = strlen( i_text ).
    " use scale factor based on default 11
    " ( don't know where defaultsetting is stored currently )
    r_width = r_width * me->size / 11.
  ENDMETHOD.