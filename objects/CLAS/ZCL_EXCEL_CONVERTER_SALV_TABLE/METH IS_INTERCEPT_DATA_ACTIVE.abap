  METHOD is_intercept_data_active.

    DATA: lr_s_type_runtime_info TYPE REF TO data.
    FIELD-SYMBOLS: <ls_type_runtime_info> TYPE any,
                   <lv_display>           TYPE any,
                   <lv_data>              TYPE any.

    rv_result = abap_false.
    TRY.
        CREATE DATA lr_s_type_runtime_info TYPE ('CL_SALV_BS_RUNTIME_INFO=>S_TYPE_RUNTIME_INFO').
        ASSIGN lr_s_type_runtime_info->* TO <ls_type_runtime_info>.
        CALL METHOD ('CL_SALV_BS_RUNTIME_INFO')=>('GET')
          RECEIVING
            value = <ls_type_runtime_info>.
        ASSIGN ('<LS_TYPE_RUNTIME_INFO>-DISPLAY') TO <lv_display>.
        CHECK sy-subrc = 0.
        ASSIGN ('<LS_TYPE_RUNTIME_INFO>-DATA') TO <lv_data>.
        CHECK sy-subrc = 0.
        IF <lv_display> = abap_false AND <lv_data> = abap_true.
          rv_result = abap_true.
        ENDIF.
      CATCH cx_sy_create_data_error cx_sy_dyn_call_error cx_salv_bs_sc_runtime_info.
        rv_result = abap_false.
    ENDTRY.

  ENDMETHOD.