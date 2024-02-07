  METHOD zif_excel_style_changer~set_font_color_rgb.

    complete_style-font-color-rgb = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  ENDMETHOD.