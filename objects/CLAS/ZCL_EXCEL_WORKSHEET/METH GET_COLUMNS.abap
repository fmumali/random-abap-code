  METHOD get_columns.

    DATA: columns TYPE TABLE OF i,
          column  TYPE i.
    FIELD-SYMBOLS:
          <sheet_cell> TYPE zexcel_s_cell_data.

    LOOP AT sheet_content ASSIGNING <sheet_cell>.
      COLLECT <sheet_cell>-cell_column INTO columns.
    ENDLOOP.

    LOOP AT columns INTO column.
      " This will create the column instance if it doesn't exist
      get_column( column ).
    ENDLOOP.

    eo_columns = me->columns.
  ENDMETHOD.                    "GET_COLUMNS