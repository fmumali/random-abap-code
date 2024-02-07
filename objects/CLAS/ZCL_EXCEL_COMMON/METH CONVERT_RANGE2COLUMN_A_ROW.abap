  METHOD convert_range2column_a_row.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          added exceptionclass
*          added errorhandling for invalid range
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#241 - error when sheetname contains "!"
*           - sheetname should be returned unescaped
*              - Stefan Schmoecker,                          2012-12-07
* changes: changed coding to support sheetnames with "!"
*          unescaping sheetname
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-07
* changes: i_range changed to clike
*          e_sheet changed to clike
*--------------------------------------------------------------------*

    DATA: lv_sheet           TYPE string,
          lv_range           TYPE string,
          lv_columnrow_start TYPE string,
          lv_columnrow_end   TYPE string,
          lv_position        TYPE i,
          lv_errormessage    TYPE string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Split input range into sheetname and Area
* 4 cases - a) input empty --> nothing to do
*         - b) sheetname existing - starts with '            example 'Sheet 1'!$B$6:$D$13
*         - c) sheetname existing - does not start with '    example Sheet1!$B$6:$D$13
*         - d) no sheetname - just area                      example $B$6:$D$13
*--------------------------------------------------------------------*
* Initialize output parameters
    CLEAR: e_column_start,
           e_column_end,
           e_row_start,
           e_row_end,
           e_sheet.

    IF i_range IS INITIAL.                                " a) input empty --> nothing to do
      RETURN.

    ELSEIF i_range(1) = `'`.                              " b) sheetname existing - starts with '
      FIND REGEX '\![^\!]*$' IN i_range MATCH OFFSET lv_position.  " Find last !
      IF sy-subrc = 0.
        lv_sheet = i_range(lv_position).
        ADD 1 TO lv_position.
        lv_range = i_range.
        SHIFT lv_range LEFT BY lv_position PLACES.
      ELSE.
        lv_errormessage = 'Invalid range'(001).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.

    ELSEIF i_range CS '!'.                                " c) sheetname existing - does not start with '
      SPLIT i_range AT '!' INTO lv_sheet lv_range.
      " begin Dennis Schaaf
      IF lv_range CP '*#REF*'.
        lv_errormessage = 'Invalid range'(001).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.
      " end Dennis Schaaf
    ELSE.                                                 " d) no sheetname - just area
      lv_range = i_range.
    ENDIF.

    REPLACE ALL OCCURRENCES OF '$' IN lv_range WITH ''.
    SPLIT lv_range AT ':' INTO lv_columnrow_start lv_columnrow_end.

    IF i_allow_1dim_range = abap_true.
      convert_columnrow2column_o_row( EXPORTING i_columnrow = lv_columnrow_start
                                      IMPORTING e_column    = e_column_start
                                                e_row       = e_row_start ).
      convert_columnrow2column_o_row( EXPORTING i_columnrow = lv_columnrow_end
                                      IMPORTING e_column    = e_column_end
                                                e_row       = e_row_end ).
    ELSE.
      convert_columnrow2column_a_row( EXPORTING i_columnrow = lv_columnrow_start
                                      IMPORTING e_column    = e_column_start
                                                e_row       = e_row_start ).
      convert_columnrow2column_a_row( EXPORTING i_columnrow = lv_columnrow_end
                                      IMPORTING e_column    = e_column_end
                                                e_row       = e_row_end ).
    ENDIF.

    IF e_column_start_int IS SUPPLIED AND e_column_start IS NOT INITIAL.
      e_column_start_int = convert_column2int( e_column_start ).
    ENDIF.
    IF e_column_end_int IS SUPPLIED AND e_column_end IS NOT INITIAL.
      e_column_end_int = convert_column2int( e_column_end ).
    ENDIF.

    e_sheet = unescape_string( lv_sheet ).                  " Return in unescaped form
  ENDMETHOD.