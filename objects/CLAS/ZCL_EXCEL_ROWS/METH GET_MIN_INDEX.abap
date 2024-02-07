  METHOD get_min_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF ep_index = 0 OR <ls_hashed_row>-row_index < ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.