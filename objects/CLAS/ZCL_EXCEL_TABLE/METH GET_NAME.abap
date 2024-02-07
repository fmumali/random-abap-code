  METHOD get_name.

    IF me->name IS INITIAL.
      me->name = zcl_excel_common=>number_to_excel_string( ip_value = me->id ).
      CONCATENATE 'table' me->name INTO me->name.
    ENDIF.

    ov_name = me->name.
  ENDMETHOD.