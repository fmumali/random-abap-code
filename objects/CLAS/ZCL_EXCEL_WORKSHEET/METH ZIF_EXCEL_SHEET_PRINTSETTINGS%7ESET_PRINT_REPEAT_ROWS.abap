  METHOD zif_excel_sheet_printsettings~set_print_repeat_rows.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    DATA:     lv_errormessage                 TYPE string.


*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    IF iv_rows_from < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF  iv_rows_from > iv_rows_to.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_row_from = iv_rows_from.
    me->print_title_row_to   = iv_rows_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_ROWS