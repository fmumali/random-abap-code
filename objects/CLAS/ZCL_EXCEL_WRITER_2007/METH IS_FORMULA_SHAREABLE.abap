  METHOD is_formula_shareable.
    DATA: lv_test_shared TYPE string.

    ep_shareable = abap_false.
    IF ip_formula NA '!'.
      lv_test_shared = zcl_excel_common=>shift_formula(
          iv_reference_formula = ip_formula
          iv_shift_cols        = 1
          iv_shift_rows        = 1 ).
      IF lv_test_shared <> ip_formula.
        ep_shareable = abap_true.
      ENDIF.
    ENDIF.
  ENDMETHOD.