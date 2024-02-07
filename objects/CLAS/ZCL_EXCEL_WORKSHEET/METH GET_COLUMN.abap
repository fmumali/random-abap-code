  METHOD get_column.

    DATA: lv_column TYPE zexcel_cell_column.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    eo_column = me->columns->get( ip_index = lv_column ).

    IF eo_column IS NOT BOUND.
      eo_column = me->add_new_column( ip_column ).
    ENDIF.

  ENDMETHOD.                    "GET_COLUMN