  METHOD get_drawings.

    DATA: lo_drawing  TYPE REF TO zcl_excel_drawing,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        r_drawings = drawings.
      WHEN zcl_excel_drawing=>type_chart.
        r_drawings = charts.
      WHEN space.
        CREATE OBJECT r_drawings
          EXPORTING
            ip_type = ''.

        lo_iterator = drawings->get_iterator( ).
        WHILE lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        ENDWHILE.
        lo_iterator = charts->get_iterator( ).
        WHILE lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        ENDWHILE.
      WHEN OTHERS.
    ENDCASE.
  ENDMETHOD.                    "GET_DRAWINGS