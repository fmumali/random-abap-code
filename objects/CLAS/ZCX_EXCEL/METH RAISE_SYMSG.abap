  METHOD raise_symsg.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        syst_at_raise = syst.
  ENDMETHOD.