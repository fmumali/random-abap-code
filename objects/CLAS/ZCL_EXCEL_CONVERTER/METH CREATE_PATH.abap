  METHOD create_path.
    DATA: l_sep    TYPE c,
          l_path   TYPE string,
          l_return TYPE i.

    CLEAR r_path.

    " Save the file
    cl_gui_frontend_services=>get_sapgui_workdir(
      CHANGING
        sapworkdir            = l_path
          EXCEPTIONS
            get_sapworkdir_failed = 1
            cntl_error            = 2
            error_no_gui          = 3
            not_supported_by_gui  = 4
           ).
    IF sy-subrc <> 0.
      CONCATENATE 'Excel_' w_fcount '.xlsx' INTO r_path.
    ELSE.
      DO.
        ADD 1 TO w_fcount.
*-obtain file separator character---------------------------------------
        CALL METHOD cl_gui_frontend_services=>get_file_separator
          CHANGING
            file_separator       = l_sep
          EXCEPTIONS
            cntl_error           = 1
            error_no_gui         = 2
            not_supported_by_gui = 3
            OTHERS               = 4.

        IF sy-subrc <> 0.
          l_sep = ''.
        ENDIF.

        CONCATENATE l_path l_sep 'Excel_' w_fcount '.xlsx' INTO r_path.

        IF cl_gui_frontend_services=>file_exist( file  = r_path ) = abap_true.
          cl_gui_frontend_services=>file_delete( EXPORTING filename = r_path
                                                 CHANGING  rc       = l_return
                                                 EXCEPTIONS OTHERS  = 1 ).
          IF sy-subrc = 0 .
            RETURN.
          ENDIF.
        ELSE.
          RETURN.
        ENDIF.
      ENDDO.
    ENDIF.

  ENDMETHOD.