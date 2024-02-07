  METHOD write_file.
    DATA: l_bytecount TYPE i,
          lt_file     TYPE solix_tab,
          l_dir       TYPE string.

    FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.

    ASSIGN wo_data->* TO <fs_data>.

    IF wo_excel IS BOUND.
      get_file( IMPORTING e_bytecount  = l_bytecount
                          et_file      = lt_file ) .
      IF i_path IS INITIAL.
        l_dir =  create_path( ) .
      ELSE.
        l_dir = i_path.
      ENDIF.
      cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = l_bytecount
                                                        filename     = l_dir
                                                        filetype     = 'BIN'
                                               CHANGING data_tab     = lt_file ).
    ENDIF.
  ENDMETHOD.