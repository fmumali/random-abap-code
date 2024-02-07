  METHOD zif_excel_converter~can_convert_object.
    DATA: lo_alv TYPE REF TO cl_gui_alv_grid.

    TRY.
        lo_alv ?= io_object.
      CATCH cx_sy_move_cast_error .
        RAISE EXCEPTION TYPE zcx_excel.
    ENDTRY.

  ENDMETHOD.