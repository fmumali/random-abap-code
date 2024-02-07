  METHOD get_height_emu_str.
    r_height = pixel2emu( size-height ).
    CONDENSE r_height NO-GAPS.
  ENDMETHOD.