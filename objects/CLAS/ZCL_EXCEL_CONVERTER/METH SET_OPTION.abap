  METHOD set_option.

    IF ws_indx-begdt IS INITIAL.
      ws_indx-begdt = sy-datum.
    ENDIF.

    ws_indx-aedat = sy-datum.
    ws_indx-usera = sy-uname.
    ws_indx-pgmid = sy-cprog.

    EXPORT p1 = is_option TO DATABASE indx(xl) FROM ws_indx ID ws_indx-srtfd.

    IF sy-subrc = 0.
      ws_option = is_option.
    ENDIF.

  ENDMETHOD.