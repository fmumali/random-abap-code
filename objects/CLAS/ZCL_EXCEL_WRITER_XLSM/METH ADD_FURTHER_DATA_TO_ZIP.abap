  METHOD add_further_data_to_zip.

    super->add_further_data_to_zip( io_zip = io_zip ).

* Add vbaProject.bin to zip
    io_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

  ENDMETHOD.