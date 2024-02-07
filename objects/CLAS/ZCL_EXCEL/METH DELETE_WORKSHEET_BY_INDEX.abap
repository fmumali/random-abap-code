  METHOD delete_worksheet_by_index.

    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lv_errormessage TYPE string.

    lo_worksheet = me->get_worksheet_by_index( iv_index ).
    IF lo_worksheet IS NOT BOUND.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.
    me->delete_worksheet( lo_worksheet ).

  ENDMETHOD.