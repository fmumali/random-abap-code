  METHOD set_style.
    me->ns_c14styleval = ip_style-c14style.
    CONDENSE me->ns_c14styleval NO-GAPS.
    me->ns_styleval = ip_style-cstyle.
    CONDENSE me->ns_styleval NO-GAPS.
  ENDMETHOD.