  METHOD zif_excel_style_changer~set_font_underline_mode.

    complete_style-font-underline_mode = value.
    complete_stylex-font-underline_mode = 'X'.

    result = me.

  ENDMETHOD.