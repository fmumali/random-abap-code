  METHOD zif_excel_style_changer~set_complete.

    complete_style = ip_complete.
    complete_stylex = ip_xcomplete.
    multiple_change_requested-complete = abap_true.
    result = me.

  ENDMETHOD.