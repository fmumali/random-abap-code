  METHOD set_width.
    TRY.
        me->width = ip_width.
        io_column = me.
      CATCH cx_sy_conversion_no_number.
        zcx_excel=>raise_text( 'Unable to interpret width as number' ).
    ENDTRY.
  ENDMETHOD.