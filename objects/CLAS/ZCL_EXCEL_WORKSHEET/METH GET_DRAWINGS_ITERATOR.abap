  METHOD get_drawings_iterator.
    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        eo_iterator = drawings->get_iterator( ).
      WHEN zcl_excel_drawing=>type_chart.
        eo_iterator = charts->get_iterator( ).
    ENDCASE.
  ENDMETHOD.                    "GET_DRAWINGS_ITERATOR