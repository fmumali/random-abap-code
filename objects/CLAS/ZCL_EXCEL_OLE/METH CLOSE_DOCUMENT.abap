  METHOD close_document.

    DATA: l_is_closed TYPE i.

    CLEAR: l_is_closed.
    IF lo_proxy IS NOT INITIAL.

* check proxy detroyed adi

      CALL METHOD lo_proxy->is_destroyed
        IMPORTING
          ret_value = l_is_closed.

* if dun detroyed yet: close -> release proxy

      IF l_is_closed IS INITIAL.
        CALL METHOD lo_proxy->close_document
          IMPORTING
            error   = lo_error
            retcode = lc_retcode.
      ENDIF.

      CALL METHOD lo_proxy->release_document
        IMPORTING
          error   = lo_error
          retcode = lc_retcode.

    ELSE.
      lc_retcode = c_oi_errors=>ret_document_not_open.
    ENDIF.

* Detroy control container

    IF lo_control IS NOT INITIAL.
      CALL METHOD lo_control->destroy_control.
    ENDIF.

    CLEAR:
      lo_spreadsheet,
      lo_proxy,
      lo_control.

  ENDMETHOD.