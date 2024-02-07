  METHOD get_style_index_in_styles.
    DATA: index TYPE i.
    DATA: lo_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_style    TYPE REF TO zcl_excel_style.

    CHECK ip_guid IS NOT INITIAL.


    lo_iterator = me->get_styles_iterator( ).
    WHILE lo_iterator->has_next( ) = 'X'.
      ADD 1 TO index.
      lo_style ?= lo_iterator->get_next( ).
      IF lo_style->get_guid( ) = ip_guid.
        ep_index = index.
        EXIT.
      ENDIF.
    ENDWHILE.

    IF ep_index IS INITIAL.
      zcx_excel=>raise_text( 'Index not found' ).
    ELSE.
      SUBTRACT 1 FROM ep_index.  " In excel list starts with "0"
    ENDIF.
  ENDMETHOD.