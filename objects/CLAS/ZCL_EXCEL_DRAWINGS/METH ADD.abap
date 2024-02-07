  METHOD add.
    DATA: lv_index TYPE i.

    drawings->add( ip_drawing ).
    lv_index = drawings->size( ).
    ip_drawing->create_media_name(
      ip_index = lv_index ).
  ENDMETHOD.