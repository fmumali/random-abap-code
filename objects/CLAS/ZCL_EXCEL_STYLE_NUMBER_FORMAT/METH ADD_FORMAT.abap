  METHOD add_format.
    DATA ls_num_format LIKE LINE OF mt_built_in_num_formats.
    ls_num_format-id                  = id.
    CREATE OBJECT ls_num_format-format.
    ls_num_format-format->format_code = code.
    INSERT ls_num_format INTO TABLE mt_built_in_num_formats.
  ENDMETHOD.