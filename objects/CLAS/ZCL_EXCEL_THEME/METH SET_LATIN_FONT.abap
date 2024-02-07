  METHOD set_latin_font.
    elements->font_scheme->modify_latin_font(
      EXPORTING
        iv_type        = iv_type
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "set_latin_font