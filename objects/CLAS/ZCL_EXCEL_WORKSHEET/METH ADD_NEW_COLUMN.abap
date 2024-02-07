  METHOD add_new_column.
    DATA: lv_column_alpha TYPE zexcel_cell_column_alpha.

    lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column ).

    CREATE OBJECT eo_column
      EXPORTING
        ip_index     = lv_column_alpha
        ip_excel     = me->excel
        ip_worksheet = me.
    columns->add( eo_column ).
  ENDMETHOD.                    "ADD_NEW_COLUMN