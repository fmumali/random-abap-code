  METHOD set_row_height.
    DATA: height TYPE f.
    TRY.
        height = ip_row_height.
        IF height <= 0.
          zcx_excel=>raise_text( 'Please supply a positive number as row-height' ).
        ENDIF.
        me->row_height = ip_row_height.
      CATCH cx_sy_conversion_no_number.
        zcx_excel=>raise_text( 'Unable to interpret ip_row_height as number' ).
    ENDTRY.
    me->custom_height = ip_custom_height.
  ENDMETHOD.