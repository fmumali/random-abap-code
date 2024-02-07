  METHOD zif_excel_style_changer~set_font_size.

    complete_style-font-size = value.
    complete_stylex-font-size = abap_true.
    single_change_requested-font-size = abap_true.
    result = me.

  ENDMETHOD.