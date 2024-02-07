  METHOD zif_excel_style_changer~set_font_family.

    complete_style-font-family = value.
    complete_stylex-font-family = 'X'.
    single_change_requested-font-family = 'X'.

    result = me.

  ENDMETHOD.