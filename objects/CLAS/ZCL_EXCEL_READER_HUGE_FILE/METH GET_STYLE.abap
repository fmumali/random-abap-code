  METHOD get_style.

    DATA: lv_tabix TYPE i,
          lo_style TYPE REF TO zcl_excel_style,
          lv_error TYPE string.

    IF gs_buffer_style-index NE iv_index.
      lv_tabix = iv_index + 1.
      READ TABLE styles INTO lo_style INDEX lv_tabix.
      IF sy-subrc NE 0.
        CONCATENATE 'Entry ' iv_index ' not found in Style Table' INTO lv_error.
        RAISE EXCEPTION TYPE lcx_not_found
          EXPORTING
            error = lv_error.
      ELSE.
        gs_buffer_style-index = iv_index.
        gs_buffer_style-guid  = lo_style->get_guid( ).
      ENDIF.
    ENDIF.

    ev_style_guid = gs_buffer_style-guid.

  ENDMETHOD.