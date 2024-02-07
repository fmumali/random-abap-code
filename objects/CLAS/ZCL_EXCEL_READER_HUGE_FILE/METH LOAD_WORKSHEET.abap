  METHOD load_worksheet.

    DATA: lo_reader TYPE REF TO if_sxml_reader.
    DATA: lx_not_found TYPE REF TO lcx_not_found.

    lo_reader = get_sxml_reader( ip_path ).

    TRY.

        read_worksheet_data( io_reader    = lo_reader
                             io_worksheet = io_worksheet ).

      CATCH lcx_not_found INTO lx_not_found.
        zcx_excel=>raise_text( lx_not_found->error ).
    ENDTRY.
  ENDMETHOD.