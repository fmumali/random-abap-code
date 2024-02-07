  METHOD provided_string_is_escaped.

    "Check if passed value is really an escaped Character
    IF value CS '_x'.
      is_escaped = abap_true.
       TRY.
          IF substring( val = value off = sy-fdpos + 6 len = 1 ) <> '_'.
            is_escaped = abap_false.
          ENDIF.
        CATCH cx_sy_range_out_of_bounds.
          is_escaped = abap_false.
      ENDTRY.
    ENDIF.
  ENDMETHOD.