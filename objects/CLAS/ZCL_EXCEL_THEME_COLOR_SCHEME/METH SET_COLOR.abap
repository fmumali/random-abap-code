  METHOD set_color.
    FIELD-SYMBOLS: <color> TYPE t_color.
    CHECK iv_type IS NOT INITIAL.
    CHECK iv_srgb IS NOT INITIAL OR  iv_syscolorname IS NOT INITIAL.
    CASE iv_type.
      WHEN c_dark1.
        ASSIGN dark1 TO <color>.
      WHEN c_dark2.
        ASSIGN dark2 TO <color>.
      WHEN c_light1.
        ASSIGN light1 TO <color>.
      WHEN c_light2.
        ASSIGN light2 TO <color>.
      WHEN c_accent1.
        ASSIGN accent1 TO <color>.
      WHEN c_accent2.
        ASSIGN accent2 TO <color>.
      WHEN c_accent3.
        ASSIGN accent3 TO <color>.
      WHEN c_accent4.
        ASSIGN accent4 TO <color>.
      WHEN c_accent5.
        ASSIGN accent5 TO <color>.
      WHEN c_accent6.
        ASSIGN accent6 TO <color>.
      WHEN c_hlink.
        ASSIGN hlink TO <color>.
      WHEN c_folhlink.
        ASSIGN folhlink TO <color>.
    ENDCASE.
    CHECK <color> IS ASSIGNED.
    CLEAR <color>.
    IF iv_srgb IS NOT INITIAL.
      <color>-srgb = iv_srgb.
    ELSE.
      <color>-syscolor-val = iv_syscolorname.
      IF iv_syscolorlast IS NOT INITIAL.
        <color>-syscolor-lastclr = iv_syscolorlast.
      ELSE.
        <color>-syscolor-lastclr = '000000'.
      ENDIF.
    ENDIF.
  ENDMETHOD.                    "set_color