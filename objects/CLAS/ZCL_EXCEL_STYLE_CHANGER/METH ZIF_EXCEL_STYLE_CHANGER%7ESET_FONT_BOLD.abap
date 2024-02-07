  METHOD zif_excel_style_changer~set_font_bold.

    complete_style-font-bold = value.
    complete_stylex-font-bold = 'X'.
    single_change_requested-font-bold = 'X'.

    result = me.

  ENDMETHOD.