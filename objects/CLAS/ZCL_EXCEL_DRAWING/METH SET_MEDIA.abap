  METHOD set_media.
    IF ip_media IS SUPPLIED.
      media = ip_media.
    ENDIF.
    media_type = ip_media_type.
    media_source = c_media_source_xstring.
    IF ip_width IS SUPPLIED.
      size-width  = ip_width.
    ENDIF.
    IF ip_height IS SUPPLIED.
      size-height = ip_height.
    ENDIF.
  ENDMETHOD.