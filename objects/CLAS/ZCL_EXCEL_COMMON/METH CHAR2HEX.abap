  METHOD char2hex.

    IF o_conv IS NOT BOUND.
      o_conv = cl_abap_conv_out_ce=>create( endian   = 'L'
                                            ignore_cerr = abap_true
                                            replacement = '#' ).
    ENDIF.

    CALL METHOD o_conv->reset( ).
    CALL METHOD o_conv->write( data = i_char ).
    r_hex+1 = o_conv->get_buffer( ). " x'65' must be x'0065'

  ENDMETHOD.