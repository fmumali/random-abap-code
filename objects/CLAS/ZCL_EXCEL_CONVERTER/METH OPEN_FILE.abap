  METHOD open_file.
    DATA: l_bytecount TYPE i,
          lt_file     TYPE solix_tab,
          l_dir       TYPE string.

    FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.

    ASSIGN wo_data->* TO <fs_data>.

    IF wo_excel IS BOUND.
      get_file( IMPORTING e_bytecount  = l_bytecount
                          et_file      = lt_file ) .

      l_dir =  create_path( ) .

      cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = l_bytecount
                                                        filename     = l_dir
                                                        filetype     = 'BIN'
                                               CHANGING data_tab     = lt_file ).
      cl_gui_frontend_services=>execute(
        EXPORTING
          document               = l_dir
        EXCEPTIONS
          cntl_error             = 1
          error_no_gui           = 2
          bad_parameter          = 3
          file_not_found         = 4
          path_not_found         = 5
          file_extension_unknown = 6
          error_execute_failed   = 7
          synchronous_failed     = 8
          not_supported_by_gui   = 9
             ).
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                   WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

    ENDIF.


  ENDMETHOD.