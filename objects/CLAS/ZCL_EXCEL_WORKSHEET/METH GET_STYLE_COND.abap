  METHOD get_style_cond.

    DATA: lo_style_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_style_cond     TYPE REF TO zcl_excel_style_cond.

    lo_style_iterator = me->get_style_cond_iterator( ).
    WHILE lo_style_iterator->has_next( ) = abap_true.
      lo_style_cond ?= lo_style_iterator->get_next( ).
      IF lo_style_cond->get_guid( ) = ip_guid.
        eo_style_cond = lo_style_cond.
        EXIT.
      ENDIF.
    ENDWHILE.

  ENDMETHOD.                    "GET_STYLE_COND