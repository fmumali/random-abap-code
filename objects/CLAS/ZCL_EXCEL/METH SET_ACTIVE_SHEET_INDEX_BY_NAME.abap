  METHOD set_active_sheet_index_by_name.

    DATA: ws_it    TYPE REF TO zcl_excel_collection_iterator,
          ws       TYPE REF TO zcl_excel_worksheet,
          lv_title TYPE zexcel_sheet_title,
          count    TYPE i VALUE 1.

    ws_it = me->worksheets->get_iterator( ).

    WHILE ws_it->has_next( ) = abap_true.
      ws ?= ws_it->get_next( ).
      lv_title = ws->get_title( ).
      IF lv_title = i_worksheet_name.
        me->worksheets->active_worksheet = count.
        EXIT.
      ENDIF.
      count = count + 1.
    ENDWHILE.

  ENDMETHOD.