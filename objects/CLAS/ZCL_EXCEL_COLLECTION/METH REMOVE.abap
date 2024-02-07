  METHOD remove .
    DATA obj TYPE REF TO object.
    LOOP AT collection INTO obj.
      IF obj = element.
        DELETE collection.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.