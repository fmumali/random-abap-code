  METHOD is_security_enabled.
    IF lockrevision EQ abap_true OR lockstructure EQ abap_true OR lockwindows EQ abap_true.
      ep_security_enabled = abap_true.
    ENDIF.
  ENDMETHOD.