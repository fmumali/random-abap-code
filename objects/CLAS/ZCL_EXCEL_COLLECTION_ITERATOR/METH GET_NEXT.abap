  METHOD get_next .
    DATA obj TYPE REF TO object.
    index = index + 1.
    object = collection->get( index ).
  ENDMETHOD.