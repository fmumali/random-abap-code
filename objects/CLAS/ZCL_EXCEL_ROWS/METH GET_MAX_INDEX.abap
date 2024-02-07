  METHOD get_max_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF <ls_hashed_row>-row_index > ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.