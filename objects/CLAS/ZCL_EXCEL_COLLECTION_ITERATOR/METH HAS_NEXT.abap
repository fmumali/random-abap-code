  METHOD has_next.
    DATA obj TYPE REF TO object.
    obj = collection->get( index + 1 ).
    has_next = boolc( obj IS NOT INITIAL ).
  ENDMETHOD.