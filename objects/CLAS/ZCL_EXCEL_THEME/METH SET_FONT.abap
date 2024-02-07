  METHOD set_font.
    elements->font_scheme->modify_font(
      EXPORTING
        iv_type     = iv_type
        iv_script   = iv_script
        iv_typeface = iv_typeface
    ).
  ENDMETHOD.                    "set_font