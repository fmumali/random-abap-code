  METHOD add_new_worksheet.

* Create default blank worksheet
    CREATE OBJECT eo_worksheet
      EXPORTING
        ip_excel = me
        ip_title = ip_title.

    worksheets->add( eo_worksheet ).
    worksheets->active_worksheet = worksheets->size( ).
  ENDMETHOD.