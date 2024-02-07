  METHOD convert_column2int.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-29
*              - ...
* changes: renaming variables to naming conventions
*          removing unused variables
*          removing commented out code that is inactive for more then half a year
*          message made to support multilinguality
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#246 - error converting lower case column names
*              - Stefan Schmoecker,                          2012-12-29
* changes: translating the correct variable to upper dase
*          adding missing exception if input is a number
*          that is out of bounds
*          adding missing exception if input contains
*          illegal characters like german umlauts
*--------------------------------------------------------------------*

    DATA: lv_column       TYPE zexcel_cell_column_alpha,
          lv_column_c     TYPE c LENGTH 10,
          lv_column_s     TYPE string,
          lv_errormessage TYPE string,                          " Can't pass '...'(abc) to exception-class
          lv_modulo       TYPE i.

*--------------------------------------------------------------------*
* This module tries to identify which column a user wants to access
* Numbers as input are just passed back, anything else will be converted
* using EXCEL nomenclatura A = 1, AA = 27, ..., XFD = 16384
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* Normalize input ( upper case , no gaps )
*--------------------------------------------------------------------*
    lv_column_c = ip_column.
    TRANSLATE lv_column_c TO UPPER CASE.                      " Fix #246
    CONDENSE lv_column_c NO-GAPS.
    IF lv_column_c EQ ''.
      MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
      zcx_excel=>raise_symsg( ).
    ENDIF.

*--------------------------------------------------------------------*
* Look up for previous succesfull cached result
*--------------------------------------------------------------------*
    IF lv_column_c = sv_prev_in2 AND sv_prev_out2 IS NOT INITIAL.
      ep_column = sv_prev_out2.
      RETURN.
    ELSE.
      CLEAR sv_prev_out2.
      sv_prev_in2 = lv_column_c.
    ENDIF.

*--------------------------------------------------------------------*
* If a number gets passed, just convert it to an integer and return
* the converted value
*--------------------------------------------------------------------*
    TRY.
        IF lv_column_c CO '1234567890 '.                      " Fix #164
          ep_column = lv_column_c.                            " Fix #164
*--------------------------------------------------------------------*
* Maximum column for EXCEL:  XFD = 16384    " if anyone has a reference for this information - please add here instead of this comment
*--------------------------------------------------------------------*
          IF ep_column > 16384 OR ep_column < 1.
            lv_errormessage = 'Index out of bounds'(004).
            zcx_excel=>raise_text( lv_errormessage ).
          ENDIF.
          RETURN.
        ENDIF.
      CATCH cx_sy_conversion_no_number.                 "#EC NO_HANDLER
        " Try the character-approach if approach via number has failed
    ENDTRY.

*--------------------------------------------------------------------*
* Raise error if unexpected characters turns up
*--------------------------------------------------------------------*
    lv_column_s = lv_column_c.
    IF lv_column_s CN sy-abcde.
      MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
      zcx_excel=>raise_symsg( ).
    ENDIF.

    DO 1 TIMES. "Because of using CHECK
*--------------------------------------------------------------------*
* Interpret input as number to base 26 with A=1, ... Z=26
* Raise error if unexpected character turns up
*--------------------------------------------------------------------*
* 1st character
*--------------------------------------------------------------------*
      lv_column = lv_column_c.
      FIND lv_column+0(1) IN sy-abcde MATCH OFFSET lv_modulo.
      lv_modulo = lv_modulo + 1.
      IF lv_modulo < 1 OR lv_modulo > 26.
        MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
        zcx_excel=>raise_symsg( ).
      ENDIF.
      ep_column = lv_modulo.                    " Leftmost digit

*--------------------------------------------------------------------*
* 2nd character if present
*--------------------------------------------------------------------*
      CHECK lv_column+1(1) IS NOT INITIAL.      " No need to continue if string ended
      FIND lv_column+1(1) IN sy-abcde MATCH OFFSET lv_modulo.
      lv_modulo = lv_modulo + 1.
      IF lv_modulo < 1 OR lv_modulo > 26.
        MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
        zcx_excel=>raise_symsg( ).
      ENDIF.
      ep_column = 26 * ep_column + lv_modulo.   " if second digit is present first digit is for 26^1

*--------------------------------------------------------------------*
* 3rd character if present
*--------------------------------------------------------------------*
      CHECK lv_column+2(1) IS NOT INITIAL.      " No need to continue if string ended
      FIND lv_column+2(1) IN sy-abcde MATCH OFFSET lv_modulo.
      lv_modulo = lv_modulo + 1.
      IF lv_modulo < 1 OR lv_modulo > 26.
        MESSAGE e800(zabap2xlsx) INTO lv_errormessage.
        zcx_excel=>raise_symsg( ).
      ENDIF.
      ep_column = 26 * ep_column + lv_modulo.   " if third digit is present first digit is for 26^2 and second digit for 26^1
    ENDDO.

*--------------------------------------------------------------------*
* Maximum column for EXCEL:  XFD = 16384    " if anyone has a reference for this information - please add here instead of this comment
*--------------------------------------------------------------------*
    IF ep_column > 16384 OR ep_column < 1.
      lv_errormessage = 'Index out of bounds'(004).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* Save succesfull output into cache
*--------------------------------------------------------------------*
    sv_prev_out2 = ep_column.

  ENDMETHOD.