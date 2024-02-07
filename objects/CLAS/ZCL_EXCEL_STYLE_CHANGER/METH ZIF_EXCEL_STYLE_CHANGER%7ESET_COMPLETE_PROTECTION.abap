  METHOD zif_excel_style_changer~set_complete_protection.

    MOVE-CORRESPONDING ip_protection  TO complete_style-protection.
    IF ip_xprotection IS SUPPLIED.
      MOVE-CORRESPONDING ip_xprotection TO complete_stylex-protection.
    ELSE.
      IF ip_protection-hidden IS NOT INITIAL.
        complete_stylex-protection-hidden = 'X'.
      ENDIF.
      IF ip_protection-locked IS NOT INITIAL.
        complete_stylex-protection-locked = 'X'.
      ENDIF.
    ENDIF.
    multiple_change_requested-protection = abap_true.

    result = me.

  ENDMETHOD.