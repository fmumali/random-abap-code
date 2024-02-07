  METHOD zif_excel_style_changer~set_font_color_indexed.

    complete_style-font-color-indexed = value.
    complete_stylex-font-color-indexed = 'X'.
    single_change_requested-font-color-indexed = 'X'.

    result = me.

  ENDMETHOD.