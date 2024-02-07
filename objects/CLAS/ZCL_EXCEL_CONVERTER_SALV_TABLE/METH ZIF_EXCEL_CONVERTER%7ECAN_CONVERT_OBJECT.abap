  METHOD zif_excel_converter~can_convert_object.

    DATA: lo_salv TYPE REF TO cl_salv_table.

    TRY.
        lo_salv ?= io_object.
      CATCH cx_sy_move_cast_error .
        RAISE EXCEPTION TYPE zcx_excel.
    ENDTRY.

  ENDMETHOD.