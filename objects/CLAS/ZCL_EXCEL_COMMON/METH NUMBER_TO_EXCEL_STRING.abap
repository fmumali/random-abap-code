  METHOD number_to_excel_string.
    DATA: lv_value_c TYPE c LENGTH 100.

    IF ip_currency IS INITIAL.
      WRITE ip_value TO lv_value_c EXPONENT 0 NO-GROUPING NO-SIGN.
    ELSE.
      WRITE ip_value TO lv_value_c EXPONENT 0 NO-GROUPING NO-SIGN CURRENCY ip_currency.
    ENDIF.
    REPLACE ALL OCCURRENCES OF ',' IN lv_value_c WITH '.'.

    ep_value = lv_value_c.
    CONDENSE ep_value.

    IF ip_value < 0.
      CONCATENATE '-' ep_value INTO ep_value.
    ELSEIF ip_value EQ 0.
      ep_value = '0'.
    ENDIF.
  ENDMETHOD.