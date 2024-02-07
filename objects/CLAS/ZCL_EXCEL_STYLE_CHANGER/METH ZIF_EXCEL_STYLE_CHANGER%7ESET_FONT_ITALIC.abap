  METHOD zif_excel_style_changer~set_font_italic.

    complete_style-font-italic = value.
    complete_stylex-font-italic = 'X'.

    result = me.

  ENDMETHOD.