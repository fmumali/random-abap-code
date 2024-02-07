  METHOD zif_excel_converter~can_convert_object.

    DATA: lo_result TYPE REF TO cl_salv_wd_result_data_table.

    TRY.
        lo_result ?= io_object.
      CATCH cx_sy_move_cast_error .
        RAISE EXCEPTION TYPE zcx_excel.
    ENDTRY.

  ENDMETHOD.