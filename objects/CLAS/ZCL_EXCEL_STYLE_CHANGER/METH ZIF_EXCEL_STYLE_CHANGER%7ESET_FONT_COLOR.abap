  METHOD zif_excel_style_changer~set_font_color.

    complete_style-font-color = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  ENDMETHOD.