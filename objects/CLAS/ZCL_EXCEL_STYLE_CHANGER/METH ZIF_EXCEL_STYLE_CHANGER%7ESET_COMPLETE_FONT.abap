  METHOD zif_excel_style_changer~set_complete_font.

    DATA: fontx TYPE zexcel_s_cstylex_font.

    IF ip_xfont IS SUPPLIED.
      fontx = ip_xfont.
    ELSE.
* Only supplied values should be used - exception: Flags bold and italic strikethrough underline
      fontx-bold = 'X'.
      fontx-italic = 'X'.
      fontx-strikethrough = 'X'.
      fontx-underline_mode = 'X'.
      CLEAR fontx-color WITH 'X'.
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_font-color
        CHANGING
          cs_xcolor = fontx-color ).
      IF ip_font-family IS NOT INITIAL.
        fontx-family = 'X'.
      ENDIF.
      IF ip_font-name IS NOT INITIAL.
        fontx-name = 'X'.
      ENDIF.
      IF ip_font-scheme IS NOT INITIAL.
        fontx-scheme = 'X'.
      ENDIF.
      IF ip_font-size IS NOT INITIAL.
        fontx-size = 'X'.
      ENDIF.
      IF ip_font-underline_mode IS NOT INITIAL.
        fontx-underline_mode = 'X'.
      ENDIF.
    ENDIF.

    complete_style-font = ip_font.
    complete_stylex-font = fontx.
    multiple_change_requested-font = abap_true.

    result = me.

  ENDMETHOD.