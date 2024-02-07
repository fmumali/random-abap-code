  METHOD discard_overlapped.
    DATA:
       lt_range TYPE tt_ranges.
    FIELD-SYMBOLS:
      <ls_range>   TYPE ts_range,
      <ls_range_2> TYPE ts_range.

    SORT mt_range BY sheet  start  stop.

    LOOP AT mt_range ASSIGNING <ls_range>.

      LOOP AT mt_range ASSIGNING <ls_range_2> WHERE sheet =  <ls_range>-sheet
                                            AND name  <> <ls_range>-name
                                            AND stop  >= <ls_range>-start
                                            AND start <  <ls_range>-start
                                            AND stop  <  <ls_range>-stop.
        EXIT.
      ENDLOOP.

      IF sy-subrc NE 0.
        APPEND <ls_range> TO lt_range.
      ENDIF.

    ENDLOOP.

    mt_range = lt_range.

    SORT mt_range BY sheet  start  stop DESCENDING.

  ENDMETHOD.