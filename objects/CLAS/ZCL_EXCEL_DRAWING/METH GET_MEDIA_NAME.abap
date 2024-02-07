  METHOD get_media_name.
    CONCATENATE media_name  `.` media_type INTO r_name.
  ENDMETHOD.