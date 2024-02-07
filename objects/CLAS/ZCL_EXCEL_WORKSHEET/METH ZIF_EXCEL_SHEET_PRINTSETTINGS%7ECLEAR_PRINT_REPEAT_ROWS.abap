  METHOD zif_excel_sheet_printsettings~clear_print_repeat_rows.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    CLEAR:  me->print_title_row_from,
            me->print_title_row_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_ROWS