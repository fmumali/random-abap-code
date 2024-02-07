  METHOD constructor.
    " Initialise values
    me->row_index    = ip_index.
    me->row_height   = -1.
    me->visible     = abap_true.
    me->outline_level  = 0.
    me->collapsed   = abap_false.

    " set row dimension as unformatted by default
    me->xf_index = 0.
    me->custom_height = abap_false.
  ENDMETHOD.