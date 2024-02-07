  METHOD get_ref.
    ev_ref = row.
    CONDENSE ev_ref.
    CONCATENATE column ev_ref INTO ev_ref.
  ENDMETHOD.