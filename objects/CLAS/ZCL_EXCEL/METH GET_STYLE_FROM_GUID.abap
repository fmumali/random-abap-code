  METHOD get_style_from_guid.

    DATA: lo_style    TYPE REF TO zcl_excel_style,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    lo_iterator = styles->get_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_style ?= lo_iterator->get_next( ).
      IF lo_style->get_guid( ) = ip_guid.
        eo_style = lo_style.
        RETURN.
      ENDIF.
    ENDWHILE.

  ENDMETHOD.