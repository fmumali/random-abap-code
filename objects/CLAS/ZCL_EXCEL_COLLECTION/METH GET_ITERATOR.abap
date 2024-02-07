  METHOD get_iterator .
    CREATE OBJECT iterator
      EXPORTING
        collection = me.
  ENDMETHOD.