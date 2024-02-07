  METHOD modify_lec_fonts.
    FIELD-SYMBOLS: <type> TYPE t_fonts,
                   <font> TYPE t_fonttype.
    CASE iv_type.
      WHEN c_minor.
        ASSIGN font_scheme-minor TO <type>.
      WHEN c_major.
        ASSIGN font_scheme-major TO <type>.
      WHEN OTHERS.
        RETURN.
    ENDCASE.
    CHECK <type> IS ASSIGNED.
    CASE iv_font_type.
      WHEN c_latin.
        ASSIGN <type>-latin TO <font>.
      WHEN c_ea.
        ASSIGN <type>-ea TO <font>.
      WHEN c_cs.
        ASSIGN <type>-cs TO <font>.
      WHEN OTHERS.
        RETURN.
    ENDCASE.
    CHECK <font> IS ASSIGNED.
    <font>-typeface = iv_typeface.
    <font>-panose = iv_panose.
    <font>-pitchfamily = iv_pitchfamily.
    <font>-charset = iv_charset.
  ENDMETHOD.                    "modify_lec_fonts