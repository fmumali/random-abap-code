  METHOD set_media_mime.

    DATA: lv_language LIKE sy-langu.

    io = ip_io.
    media_source = c_media_source_mime.
    size-width  = ip_width.
    size-height = ip_height.

    lv_language = sy-langu.
    cl_wb_mime_repository=>load_mime( EXPORTING
                                        io        = ip_io
                                      IMPORTING
                                        filename  = media_name
                                        "mimetype = media_type
                                      CHANGING
                                        language  = lv_language  ).

    SPLIT media_name AT '.' INTO media_name media_type.

  ENDMETHOD.