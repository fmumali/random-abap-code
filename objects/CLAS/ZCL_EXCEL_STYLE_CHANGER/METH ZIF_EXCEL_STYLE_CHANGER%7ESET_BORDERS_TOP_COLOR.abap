  METHOD zif_excel_style_changer~set_borders_top_color.

    complete_style-borders-top-border_color = value.
    complete_stylex-borders-top-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.