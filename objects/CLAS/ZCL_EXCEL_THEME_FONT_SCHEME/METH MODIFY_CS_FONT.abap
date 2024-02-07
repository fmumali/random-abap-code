  METHOD modify_cs_font.
    modify_lec_fonts(
      EXPORTING
        iv_type        = iv_type
        iv_font_type   = c_cs
        iv_typeface    = iv_typeface
        iv_panose      = iv_panose
        iv_pitchfamily = iv_pitchfamily
        iv_charset     = iv_charset
    ).
  ENDMETHOD.                    "modify_latin_font