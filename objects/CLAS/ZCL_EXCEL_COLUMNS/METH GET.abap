  METHOD get.
    FIELD-SYMBOLS: <ls_hashed_column> TYPE mty_s_hashed_column.

    READ TABLE columns_hashed WITH KEY column_index = ip_index ASSIGNING <ls_hashed_column>.
    IF sy-subrc = 0.
      eo_column = <ls_hashed_column>-column.
    ENDIF.
  ENDMETHOD.