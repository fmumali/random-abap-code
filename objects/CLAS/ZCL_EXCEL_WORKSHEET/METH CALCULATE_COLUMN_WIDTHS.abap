  METHOD calculate_column_widths.
    TYPES:
      BEGIN OF t_auto_size,
        col_index TYPE int4,
        width     TYPE f,
      END   OF t_auto_size.
    TYPES: tt_auto_size TYPE TABLE OF t_auto_size.

    DATA: lo_column_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_column          TYPE REF TO zcl_excel_column.

    DATA: auto_size   TYPE flag.
    DATA: auto_sizes  TYPE tt_auto_size.
    DATA: count       TYPE int4.
    DATA: highest_row TYPE int4.
    DATA: width       TYPE f.

    FIELD-SYMBOLS: <auto_size>        LIKE LINE OF auto_sizes.

    lo_column_iterator = me->get_columns_iterator( ).
    WHILE lo_column_iterator->has_next( ) = abap_true.
      lo_column ?= lo_column_iterator->get_next( ).
      auto_size = lo_column->get_auto_size( ).
      IF auto_size = abap_true.
        APPEND INITIAL LINE TO auto_sizes ASSIGNING <auto_size>.
        <auto_size>-col_index = lo_column->get_column_index( ).
        <auto_size>-width     = -1.
      ENDIF.
    ENDWHILE.

    " There is only something to do if there are some auto-size columns
    IF NOT auto_sizes IS INITIAL.
      highest_row = me->get_highest_row( ).
      LOOP AT auto_sizes ASSIGNING <auto_size>.
        count = 1.
        WHILE count <= highest_row.
* Do not check merged cells
          IF is_cell_merged(
              ip_column    = <auto_size>-col_index
              ip_row       = count ) = abap_false.
            width = calculate_cell_width( ip_column = <auto_size>-col_index     " issue #155 - less restrictive typing for ip_column
                                          ip_row    = count ).
            IF width > <auto_size>-width.
              <auto_size>-width = width.
            ENDIF.
          ENDIF.
          count = count + 1.
        ENDWHILE.
        lo_column = me->get_column( <auto_size>-col_index ). " issue #155 - less restrictive typing for ip_column
        lo_column->set_width( <auto_size>-width ).
      ENDLOOP.
    ENDIF.

  ENDMETHOD.                    "CALCULATE_COLUMN_WIDTHS