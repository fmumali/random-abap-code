  METHOD delete_row_outline.

    DELETE me->mt_row_outlines WHERE row_from = iv_row_from
                                 AND row_to   = iv_row_to.
    IF sy-subrc <> 0.  " didn't find outline that was to be deleted
      zcx_excel=>raise_text( 'Row outline to be deleted does not exist' ).
    ENDIF.

  ENDMETHOD.                    "DELETE_ROW_OUTLINE