  METHOD class_constructor.
    DATA: ls_objects TYPE ts_alv_types.
    DATA: ls_option TYPE zexcel_s_converter_option,
          l_uname   TYPE sy-uname.

    GET PARAMETER ID 'ZUS' FIELD l_uname.
    IF l_uname IS INITIAL OR l_uname = space.
      l_uname = sy-uname.
    ENDIF.

    get_alv_converters( ).

    CONCATENATE 'EXCEL_' sy-uname INTO ws_indx-srtfd.

    IMPORT p1 = ls_option FROM DATABASE indx(xl) TO ws_indx ID ws_indx-srtfd.

    IF sy-subrc = 0.
      ws_option = ls_option.
    ELSE.
      init_option( ) .
    ENDIF.

  ENDMETHOD.