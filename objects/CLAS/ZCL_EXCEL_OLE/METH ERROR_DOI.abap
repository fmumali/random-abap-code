  METHOD error_doi.

    IF lc_retcode NE c_oi_errors=>ret_ok.
      close_document( ).
      CALL METHOD lo_error->raise_message
        EXPORTING
          type = 'E'.
      CLEAR: lo_error.
    ENDIF.

  ENDMETHOD.