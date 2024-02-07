  METHOD zif_excel_sheet_printsettings~clear_print_repeat_columns.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    CLEAR:  me->print_title_col_from,
            me->print_title_col_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_COLUMNS