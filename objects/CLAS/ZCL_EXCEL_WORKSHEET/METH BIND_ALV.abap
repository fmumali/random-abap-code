  METHOD bind_alv.
    DATA: lo_converter TYPE REF TO zcl_excel_converter.

    CREATE OBJECT lo_converter.

    TRY.
        lo_converter->convert(
          EXPORTING
            io_alv         = io_alv
            it_table       = it_table
            i_row_int      = i_top
            i_column_int   = i_left
            i_table        = i_table
            i_style_table  = table_style
            io_worksheet   = me
          CHANGING
            co_excel       = excel ).
      CATCH zcx_excel .
    ENDTRY.

  ENDMETHOD.                    "BIND_ALV