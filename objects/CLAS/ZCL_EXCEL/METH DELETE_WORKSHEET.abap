  METHOD delete_worksheet.

    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          l_size          TYPE i,
          lv_errormessage TYPE string.

    l_size = get_worksheets_size( ).
    IF l_size = 1.  " Only 1 worksheet left --> check whether this is the worksheet to be deleted
      lo_worksheet = me->get_worksheet_by_index( 1 ).
      IF lo_worksheet = io_worksheet.
        lv_errormessage = 'Deleting last remaining worksheet is not allowed'(002).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.
    ENDIF.

    me->worksheets->remove( io_worksheet ).

  ENDMETHOD.