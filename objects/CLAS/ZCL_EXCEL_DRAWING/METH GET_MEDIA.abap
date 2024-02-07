  METHOD get_media.

    DATA: lv_language LIKE sy-langu.
    DATA: lt_bin_mime TYPE sdokcntbins.
    DATA: lt_mime          TYPE tsfmime,
          lv_filesize      TYPE i,
          lv_filesizec(10).

    CASE media_source.
      WHEN c_media_source_xstring.
        r_media = media.
      WHEN c_media_source_www.
        CALL FUNCTION 'WWWDATA_IMPORT'
          EXPORTING
            key    = media_key_www
          TABLES
            mime   = lt_mime
          EXCEPTIONS
            OTHERS = 1.

        CALL FUNCTION 'WWWPARAMS_READ'
          EXPORTING
            relid = media_key_www-relid
            objid = media_key_www-objid
            name  = 'filesize'
          IMPORTING
            value = lv_filesizec.

        lv_filesize = lv_filesizec.
        CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
          EXPORTING
            input_length = lv_filesize
          IMPORTING
            buffer       = r_media
          TABLES
            binary_tab   = lt_mime
          EXCEPTIONS
            failed       = 1
            OTHERS       = 2.
      WHEN c_media_source_mime.
        lv_language = sy-langu.
        cl_wb_mime_repository=>load_mime( EXPORTING
                                            io        = me->io
                                          IMPORTING
                                            filesize  = lv_filesize
                                            bin_data  = lt_bin_mime
                                          CHANGING
                                            language  = lv_language ).

        CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
          EXPORTING
            input_length = lv_filesize
          IMPORTING
            buffer       = r_media
          TABLES
            binary_tab   = lt_bin_mime
          EXCEPTIONS
            failed       = 1
            OTHERS       = 2.
    ENDCASE.
  ENDMETHOD.