  METHOD get.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    READ TABLE rows_hashed WITH KEY row_index = ip_index ASSIGNING <ls_hashed_row>.
    IF sy-subrc = 0.
      eo_row = <ls_hashed_row>-row.
    ENDIF.
  ENDMETHOD.                    "GET