  METHOD get_drawings_iterator.

    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        eo_iterator = me->drawings->get_iterator( ).
      WHEN zcl_excel_drawing=>type_chart.
        eo_iterator = me->charts->get_iterator( ).
      WHEN OTHERS.
    ENDCASE.

  ENDMETHOD.