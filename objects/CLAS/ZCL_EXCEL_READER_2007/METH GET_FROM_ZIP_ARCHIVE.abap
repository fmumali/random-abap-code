  METHOD get_from_zip_archive.

    ASSERT zip IS BOUND. " zip object has to exist at this point

    r_content = zip->read(  i_filename ).

  ENDMETHOD.