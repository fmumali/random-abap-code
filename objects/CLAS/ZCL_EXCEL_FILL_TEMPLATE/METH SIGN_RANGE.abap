  METHOD sign_range.

    DATA: lv_tabix TYPE i.
    FIELD-SYMBOLS:
      <ls_range>   TYPE ts_range,
      <ls_range_2> TYPE ts_range.

    LOOP AT mt_range ASSIGNING <ls_range>.
      <ls_range>-id = sy-tabix.
    ENDLOOP.

    LOOP AT mt_range ASSIGNING  <ls_range>.
      lv_tabix = sy-tabix + 1.

      LOOP AT mt_range ASSIGNING  <ls_range_2>
                                  FROM lv_tabix
                                  WHERE sheet = <ls_range>-sheet.

        IF <ls_range_2>-start >= <ls_range>-start AND <ls_range_2>-stop <= <ls_range>-stop.
          <ls_range_2>-parent = <ls_range>-id.
        ENDIF.

      ENDLOOP.

    ENDLOOP.

    SORT mt_range BY id DESCENDING.
  ENDMETHOD.