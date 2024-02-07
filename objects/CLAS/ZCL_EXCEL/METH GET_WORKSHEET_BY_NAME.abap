  METHOD get_worksheet_by_name.

    DATA: lv_index TYPE zexcel_active_worksheet,
          l_size   TYPE i.

    l_size = get_worksheets_size( ).

    DO l_size TIMES.
      lv_index = sy-index.
      eo_worksheet = me->worksheets->get( lv_index ).
      IF eo_worksheet->get_title( ) = ip_sheet_name.
        RETURN.
      ENDIF.
    ENDDO.

    CLEAR eo_worksheet.

  ENDMETHOD.