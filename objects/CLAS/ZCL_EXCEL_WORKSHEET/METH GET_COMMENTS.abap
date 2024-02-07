  METHOD get_comments.
    DATA: lo_comment  TYPE REF TO zcl_excel_comment,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    CREATE OBJECT r_comments.

    lo_iterator = comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_comment ?= lo_iterator->get_next( ).
      r_comments->include( lo_comment ).
    ENDWHILE.

  ENDMETHOD.                    "get_comments