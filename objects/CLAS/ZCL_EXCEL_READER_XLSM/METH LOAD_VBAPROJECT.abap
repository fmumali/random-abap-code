  METHOD load_vbaproject.

    DATA lv_content TYPE xstring.

    lv_content = me->get_from_zip_archive( ip_path ).

    ip_excel->zif_excel_book_vba_project~set_vbaproject( lv_content ).

  ENDMETHOD.