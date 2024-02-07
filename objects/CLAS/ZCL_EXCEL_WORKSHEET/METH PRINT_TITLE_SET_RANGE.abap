  METHOD print_title_set_range.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


    DATA: lo_range_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lo_range                  TYPE REF TO zcl_excel_range,
          lv_repeat_range_sheetname TYPE string,
          lv_repeat_range_col       TYPE string,
          lv_row_char_from          TYPE char10,
          lv_row_char_to            TYPE char10,
          lv_repeat_range_row       TYPE string,
          lv_repeat_range           TYPE string.


*--------------------------------------------------------------------*
* Get range that represents printarea
* if non-existant, create it
*--------------------------------------------------------------------*
    lo_range_iterator = me->get_ranges_iterator( ).
    WHILE lo_range_iterator->has_next( ) = abap_true.

      lo_range ?= lo_range_iterator->get_next( ).
      IF lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
        EXIT.  " Found it
      ENDIF.
      CLEAR lo_range.

    ENDWHILE.


    IF me->print_title_col_from IS INITIAL AND
       me->print_title_row_from IS INITIAL.
*--------------------------------------------------------------------*
* No print titles are present,
*--------------------------------------------------------------------*
      IF lo_range IS BOUND.
        me->ranges->remove( lo_range ).
      ENDIF.
    ELSE.
*--------------------------------------------------------------------*
* Print titles are present,
*--------------------------------------------------------------------*
      IF lo_range IS NOT BOUND.
        lo_range =  me->add_new_range( ).
        lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
      ENDIF.

      lv_repeat_range_sheetname = me->get_title( ).
      lv_repeat_range_sheetname = zcl_excel_common=>escape_string( lv_repeat_range_sheetname ).

*--------------------------------------------------------------------*
* Repeat-columns
*--------------------------------------------------------------------*
      IF me->print_title_col_from IS NOT INITIAL.
        CONCATENATE lv_repeat_range_sheetname
                    '!$' me->print_title_col_from
                    ':$' me->print_title_col_to
            INTO lv_repeat_range_col.
      ENDIF.

*--------------------------------------------------------------------*
* Repeat-rows
*--------------------------------------------------------------------*
      IF me->print_title_row_from IS NOT INITIAL.
        lv_row_char_from = me->print_title_row_from.
        lv_row_char_to   = me->print_title_row_to.
        CONCATENATE '!$' lv_row_char_from
                    ':$' lv_row_char_to
            INTO lv_repeat_range_row.
        CONDENSE lv_repeat_range_row NO-GAPS.
        CONCATENATE lv_repeat_range_sheetname
                    lv_repeat_range_row
            INTO lv_repeat_range_row.
      ENDIF.

*--------------------------------------------------------------------*
* Concatenate repeat-rows and columns
*--------------------------------------------------------------------*
      IF lv_repeat_range_col IS INITIAL.
        lv_repeat_range = lv_repeat_range_row.
      ELSEIF lv_repeat_range_row IS INITIAL.
        lv_repeat_range = lv_repeat_range_col.
      ELSE.
        CONCATENATE lv_repeat_range_col lv_repeat_range_row
            INTO lv_repeat_range SEPARATED BY ','.
      ENDIF.


      lo_range->set_range_value( lv_repeat_range ).
    ENDIF.



  ENDMETHOD.                    "PRINT_TITLE_SET_RANGE