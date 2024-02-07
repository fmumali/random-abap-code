  METHOD get_next_table_id.
    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lv_tables_count TYPE i.

    lo_iterator = me->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lv_tables_count = lo_worksheet->get_tables_size( ).
      ADD lv_tables_count TO ep_id.

    ENDWHILE.

    ADD 1 TO ep_id.

  ENDMETHOD.