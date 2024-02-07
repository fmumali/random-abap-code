  METHOD constructor.

    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).      " ins issue #379 - replacement for outdated function call

    IF ip_title IS NOT INITIAL.
      title = ip_title.
    ELSE.
      title = me->guid.
    ENDIF.

    me->type = ip_type.

* inizialize dimension range
    anchor = anchor_one_cell.
    from_loc-col = 1.
    from_loc-row = 1.
  ENDMETHOD.