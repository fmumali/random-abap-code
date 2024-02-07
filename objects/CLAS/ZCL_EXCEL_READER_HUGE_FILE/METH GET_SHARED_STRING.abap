  METHOD get_shared_string.
    DATA: lv_tabix TYPE i,
          lv_error TYPE string.
    FIELD-SYMBOLS: <ls_shared_string> TYPE t_shared_string.
    lv_tabix = iv_index + 1.
    READ TABLE shared_strings ASSIGNING <ls_shared_string> INDEX lv_tabix.
    IF sy-subrc NE 0.
      CONCATENATE 'Entry ' iv_index ' not found in Shared String Table' INTO lv_error.
      RAISE EXCEPTION TYPE lcx_not_found
        EXPORTING
          error = lv_error.
    ENDIF.
    ev_value = <ls_shared_string>-value.
  ENDMETHOD.