  METHOD zif_excel_style_changer~set_borders_left_color_tint.

    complete_style-borders-left-border_color-tint = value.
    complete_stylex-borders-left-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.