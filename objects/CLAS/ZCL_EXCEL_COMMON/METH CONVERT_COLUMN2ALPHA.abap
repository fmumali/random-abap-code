  METHOD convert_column2alpha.

    DATA: lv_uccpi  TYPE i,
          lv_text   TYPE c LENGTH 2,
          lv_module TYPE int4,
          lv_column TYPE zexcel_cell_column.

* Propagate zcx_excel if error occurs           " issue #155 - less restrictive typing for ip_column
    lv_column = convert_column2int( ip_column ).  " issue #155 - less restrictive typing for ip_column

*--------------------------------------------------------------------*
* Check whether column is in allowed range for EXCEL to handle ( 1-16384 )
*--------------------------------------------------------------------*
    IF   lv_column > 16384
      OR lv_column < 1.
      zcx_excel=>raise_text( 'Index out of bounds' ).
    ENDIF.

*--------------------------------------------------------------------*
* Look up for previous succesfull cached result
*--------------------------------------------------------------------*
    IF lv_column = sv_prev_in1 AND sv_prev_out1 IS NOT INITIAL.
      ep_column = sv_prev_out1.
      RETURN.
    ELSE.
      CLEAR sv_prev_out1.
      sv_prev_in1 = lv_column.
    ENDIF.

*--------------------------------------------------------------------*
* Build alpha representation of column
*--------------------------------------------------------------------*
    WHILE lv_column GT 0.

      lv_module = ( lv_column - 1 ) MOD 26.
      lv_uccpi  = 65 + lv_module.

      lv_column = ( lv_column - lv_module ) / 26.

      lv_text   = cl_abap_conv_in_ce=>uccpi( lv_uccpi ).
      CONCATENATE lv_text ep_column INTO ep_column.

    ENDWHILE.

*--------------------------------------------------------------------*
* Save succesfull output into cache
*--------------------------------------------------------------------*
    sv_prev_out1 = ep_column.

  ENDMETHOD.