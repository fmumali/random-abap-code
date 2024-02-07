  METHOD set_name.
    IF strlen( iv_name ) > 50.
      name = iv_name(50).
    ELSE.
      name = iv_name.
    ENDIF.
  ENDMETHOD.                    "set_name