  METHOD add_new_row.
    CREATE OBJECT eo_row
      EXPORTING
        ip_index = ip_row.
    rows->add( eo_row ).
  ENDMETHOD.                    "ADD_NEW_ROW