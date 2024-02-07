  METHOD add.
    DATA: lv_index TYPE i.

    comments->add( ip_comment ).
    lv_index = comments->size( ).

  ENDMETHOD.