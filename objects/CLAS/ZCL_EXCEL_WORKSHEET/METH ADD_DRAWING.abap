  METHOD add_drawing.
    CASE ip_drawing->get_type( ).
      WHEN zcl_excel_drawing=>type_image.
        drawings->include( ip_drawing ).
      WHEN zcl_excel_drawing=>type_chart.
        charts->include( ip_drawing ).
    ENDCASE.
  ENDMETHOD.                    "ADD_DRAWING