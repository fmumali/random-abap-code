  METHOD zif_excel_style_changer~set_borders_left_color.

    complete_style-borders-left-border_color = value.
    complete_stylex-borders-left-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.