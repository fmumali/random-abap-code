  METHOD get_column_filter.

    DATA: ls_filter LIKE LINE OF me->mt_filters.

    READ TABLE me->mt_filters REFERENCE INTO rr_filter WITH TABLE KEY column = i_column.
    IF sy-subrc <> 0.
      ls_filter-column = i_column.
      INSERT ls_filter INTO TABLE me->mt_filters REFERENCE INTO rr_filter.
    ENDIF.

  ENDMETHOD.