  METHOD convert.

    IF is_option IS SUPPLIED.
      ws_option = is_option.
    ENDIF.

    execute_converter( EXPORTING io_object   = io_alv
                                 it_table    = it_table ) .

    IF io_worksheet IS SUPPLIED AND io_worksheet IS BOUND.
      wo_worksheet = io_worksheet.
    ENDIF.
    IF co_excel IS SUPPLIED.
      IF co_excel IS NOT BOUND.
        CREATE OBJECT co_excel.
        co_excel->zif_excel_book_properties~creator = sy-uname.
      ENDIF.
      wo_excel = co_excel.
    ENDIF.

* Move table to data object and clean it up
    IF wt_fieldcatalog IS NOT INITIAL.
      create_table( ).
    ELSE.
      wo_data = wo_table .
    ENDIF.

    IF wo_excel IS NOT BOUND.
      CREATE OBJECT wo_excel.
      wo_excel->zif_excel_book_properties~creator = sy-uname.
    ENDIF.
    IF wo_worksheet IS NOT BOUND.
      " Get active sheet
      wo_worksheet = wo_excel->get_active_worksheet( ).
      wo_worksheet->set_title( ip_title = 'Sheet1'(001) ).
    ENDIF.

    IF i_row_int <= 0.
      w_row_int = 1.
    ELSE.
      w_row_int = i_row_int.
    ENDIF.
    IF i_column_int <= 0.
      w_col_int = 1.
    ELSE.
      w_col_int = i_column_int.
    ENDIF.

    create_worksheet( i_table       = i_table
                      i_style_table = i_style_table ) .

  ENDMETHOD.