  METHOD split_file.

    DATA: lt_hlp TYPE TABLE OF text255,
          ls_hlp TYPE text255.

    DATA: lf_ext(10)     TYPE c,
          lf_dot_ext(10) TYPE c.
    DATA: lf_anz TYPE i,
          lf_len TYPE i.
** ---------------------------------------------------------------------

    CLEAR: lt_hlp,
           ep_file,
           ep_extension,
           ep_dotextension.

** Split the whole file at '.'
    SPLIT ip_file AT '.' INTO TABLE lt_hlp.

** get the extenstion from the last line of table
    DESCRIBE TABLE lt_hlp LINES lf_anz.
    IF lf_anz <= 1.
      ep_file = ip_file.
      RETURN.
    ENDIF.

    READ TABLE lt_hlp INTO ls_hlp INDEX lf_anz.
    ep_extension = ls_hlp.
    lf_ext =  ls_hlp.
    IF NOT lf_ext IS INITIAL.
      CONCATENATE '.' lf_ext INTO lf_dot_ext.
    ENDIF.
    ep_dotextension = lf_dot_ext.

** get only the filename
    lf_len = strlen( ip_file ) - strlen( lf_dot_ext ).
    IF lf_len > 0.
      ep_file = ip_file(lf_len).
    ENDIF.

  ENDMETHOD.