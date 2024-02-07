  METHOD add_new_drawing.
* Create default blank worksheet
    CREATE OBJECT eo_drawing
      EXPORTING
        ip_type  = ip_type
        ip_title = ip_title.

    CASE ip_type.
      WHEN 'image'.
        drawings->add( eo_drawing ).
      WHEN 'hd_ft'.
        drawings->add( eo_drawing ).
      WHEN 'chart'.
        charts->add( eo_drawing ).
    ENDCASE.
  ENDMETHOD.