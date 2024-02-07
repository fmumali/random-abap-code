  METHOD read_from_local_file.
    DATA: lv_filelength   TYPE i,
          lt_binary_data  TYPE STANDARD TABLE OF x255 WITH NON-UNIQUE DEFAULT KEY,
          ls_binary_data  LIKE LINE OF lt_binary_data,
          lv_filename     TYPE string,
          lv_errormessage TYPE string.

    lv_filename = i_filename.

    cl_gui_frontend_services=>gui_upload( EXPORTING
                                            filename                = lv_filename
                                            filetype                = 'BIN'         " We are basically working with zipped directories --> force binary read
                                          IMPORTING
                                            filelength              = lv_filelength
                                          CHANGING
                                            data_tab                = lt_binary_data
                                          EXCEPTIONS
                                            file_open_error         = 1
                                            file_read_error         = 2
                                            no_batch                = 3
                                            gui_refuse_filetransfer = 4
                                            invalid_type            = 5
                                            no_authority            = 6
                                            unknown_error           = 7
                                            bad_data_format         = 8
                                            header_not_allowed      = 9
                                            separator_not_allowed   = 10
                                            header_too_long         = 11
                                            unknown_dp_error        = 12
                                            access_denied           = 13
                                            dp_out_of_memory        = 14
                                            disk_full               = 15
                                            dp_timeout              = 16
                                            not_supported_by_gui    = 17
                                            error_no_gui            = 18
                                            OTHERS                  = 19 ).
    IF sy-subrc <> 0.
      lv_errormessage = 'A problem occured when reading the file'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* Binary data needs to be provided as XSTRING for further processing
*--------------------------------------------------------------------*
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_filelength
      IMPORTING
        buffer       = r_excel_data
      TABLES
        binary_tab   = lt_binary_data.

  ENDMETHOD.