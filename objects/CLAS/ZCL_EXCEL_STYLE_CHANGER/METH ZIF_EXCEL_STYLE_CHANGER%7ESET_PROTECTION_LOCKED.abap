  METHOD zif_excel_style_changer~set_protection_locked.

    complete_style-protection-locked = value.
    complete_stylex-protection-locked = 'X'.

    result = me.

  ENDMETHOD.