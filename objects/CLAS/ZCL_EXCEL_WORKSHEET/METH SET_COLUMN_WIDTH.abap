  METHOD set_column_width.
    DATA: lo_column  TYPE REF TO zcl_excel_column.
    DATA: width             TYPE f.

    lo_column = me->get_column( ip_column ).

* if a fix size is supplied use this
    IF ip_width_fix IS SUPPLIED.
      TRY.
          width = ip_width_fix.
          IF width <= 0.
            zcx_excel=>raise_text( 'Please supply a positive number as column-width' ).
          ENDIF.
          lo_column->set_width( width ).
          RETURN.
        CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
          zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
      ENDTRY.
    ENDIF.

* If we get down to here, we have to use whatever is found in autosize.
    lo_column->set_auto_size( ip_width_autosize ).


  ENDMETHOD.                    "SET_COLUMN_WIDTH