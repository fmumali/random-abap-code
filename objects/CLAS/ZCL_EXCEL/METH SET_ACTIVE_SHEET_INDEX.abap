  METHOD set_active_sheet_index.
    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lv_errormessage TYPE string.

*--------------------------------------------------------------------*
* Check whether worksheet exists
*--------------------------------------------------------------------*
    lo_worksheet = me->get_worksheet_by_index( i_active_worksheet ).
    IF lo_worksheet IS NOT BOUND.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    me->worksheets->active_worksheet = i_active_worksheet.

  ENDMETHOD.