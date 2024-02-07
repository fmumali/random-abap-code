  METHOD raise_text.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = iv_text.
  ENDMETHOD.