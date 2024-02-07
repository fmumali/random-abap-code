  METHOD zif_excel_style_changer~set_borders_top_color_tint.

    complete_style-borders-top-border_color-tint = value.
    complete_stylex-borders-top-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.