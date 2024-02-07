  METHOD get_default_row.
    IF me->row_default IS NOT BOUND.
      CREATE OBJECT me->row_default.
    ENDIF.

    eo_row = me->row_default.
  ENDMETHOD.                    "GET_DEFAULT_ROW