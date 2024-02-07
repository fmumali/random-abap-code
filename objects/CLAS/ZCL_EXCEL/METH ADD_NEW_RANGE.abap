  METHOD add_new_range.
* Create default blank range
    CREATE OBJECT eo_range.
    ranges->add( eo_range ).
  ENDMETHOD.