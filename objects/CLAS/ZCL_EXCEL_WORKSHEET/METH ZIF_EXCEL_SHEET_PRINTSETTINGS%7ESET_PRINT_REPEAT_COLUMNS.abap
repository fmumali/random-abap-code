  METHOD zif_excel_sheet_printsettings~set_print_repeat_columns.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    DATA: lv_col_from_int TYPE i,
          lv_col_to_int   TYPE i,
          lv_errormessage TYPE string.


    lv_col_from_int = zcl_excel_common=>convert_column2int( iv_columns_from ).
    lv_col_to_int   = zcl_excel_common=>convert_column2int( iv_columns_to ).

*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    IF lv_col_from_int < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF  lv_col_from_int > lv_col_to_int.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_col_from = iv_columns_from.
    me->print_title_col_to   = iv_columns_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).

  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_COLUMNS