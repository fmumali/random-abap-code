  METHOD set_media_www.
    DATA: lv_value(20).

    media_key_www = ip_key.
    media_source = c_media_source_www.

    CALL FUNCTION 'WWWPARAMS_READ'
      EXPORTING
        relid = media_key_www-relid
        objid = media_key_www-objid
        name  = 'fileextension'
      IMPORTING
        value = lv_value.
    media_type = lv_value.
    SHIFT media_type LEFT DELETING LEADING '.'.

    size-width  = ip_width.
    size-height = ip_height.
  ENDMETHOD.