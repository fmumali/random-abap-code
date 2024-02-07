  METHOD set_row_height.
    DATA: lo_row  TYPE REF TO zcl_excel_row.
    DATA: height  TYPE f.

    lo_row = me->get_row( ip_row ).

* if a fix size is supplied use this
    TRY.
        height = ip_height_fix.
        lo_row->set_row_height( height ).
        RETURN.
      CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
        zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
    ENDTRY.

  ENDMETHOD.                    "SET_ROW_HEIGHT