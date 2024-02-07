  METHOD get_default_column.
    IF me->column_default IS NOT BOUND.
      CREATE OBJECT me->column_default
        EXPORTING
          ip_index     = 'A'         " ????
          ip_worksheet = me
          ip_excel     = me->excel.
    ENDIF.

    eo_column = me->column_default.
  ENDMETHOD.                    "GET_DEFAULT_COLUMN