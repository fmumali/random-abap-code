  METHOD get_width_emu_str.
    r_width = pixel2emu( size-width ).
    CONDENSE r_width NO-GAPS.
  ENDMETHOD.