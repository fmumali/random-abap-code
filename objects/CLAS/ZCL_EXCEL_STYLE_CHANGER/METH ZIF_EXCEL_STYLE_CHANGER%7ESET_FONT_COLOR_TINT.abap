  METHOD zif_excel_style_changer~set_font_color_tint.

    complete_style-font-color-tint = value.
    complete_stylex-font-color-tint = 'X'.
    single_change_requested-font-color-tint = 'X'.

    result = me.

  ENDMETHOD.