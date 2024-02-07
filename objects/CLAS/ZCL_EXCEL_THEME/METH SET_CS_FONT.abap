  METHOD set_cs_font.
    elements->font_scheme->modify_cs_font(
      EXPORTING
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "set_cs_font