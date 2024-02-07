  METHOD zif_excel_style_changer~set_number_format.

    complete_style-number_format-format_code = value.
    complete_stylex-number_format-format_code = abap_true.
    single_change_requested-number_format-format_code = abap_true.
    result = me.

  ENDMETHOD.