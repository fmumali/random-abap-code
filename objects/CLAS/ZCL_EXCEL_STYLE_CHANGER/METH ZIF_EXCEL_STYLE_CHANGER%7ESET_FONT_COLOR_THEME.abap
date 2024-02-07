  METHOD zif_excel_style_changer~set_font_color_theme.

    complete_style-font-color-theme = value.
    complete_stylex-font-color-theme = 'X'.
    single_change_requested-font-color-theme = 'X'.

    result = me.

  ENDMETHOD.