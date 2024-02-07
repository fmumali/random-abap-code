  METHOD clear_initial_colorxfields.

    IF is_color-rgb IS INITIAL.
      CLEAR cs_xcolor-rgb.
    ENDIF.
    IF is_color-indexed IS INITIAL.
      CLEAR cs_xcolor-indexed.
    ENDIF.
    IF is_color-theme IS INITIAL.
      CLEAR cs_xcolor-theme.
    ENDIF.
    IF is_color-tint IS INITIAL.
      CLEAR cs_xcolor-tint.
    ENDIF.

  ENDMETHOD.