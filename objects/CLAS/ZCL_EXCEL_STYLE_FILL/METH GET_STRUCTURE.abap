  METHOD get_structure.
    es_fill-rotation  = me->rotation.
    es_fill-filltype  = me->filltype.
    es_fill-fgcolor   = me->fgcolor.
    es_fill-bgcolor   = me->bgcolor.
    me->build_gradient( ).
    es_fill-gradtype = me->gradtype.
  ENDMETHOD.                    "GET_STRUCTURE