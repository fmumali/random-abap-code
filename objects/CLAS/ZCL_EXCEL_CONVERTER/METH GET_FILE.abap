  METHOD get_file.
    DATA: lo_excel_writer         TYPE REF TO zif_excel_writer.

    DATA: ls_seoclass TYPE seoclass.


    IF wo_excel IS BOUND.
      CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
      e_file = lo_excel_writer->write_file( wo_excel ).

      SELECT SINGLE * INTO ls_seoclass
        FROM seoclass
        WHERE clsname = 'CL_BCS_CONVERT'.

      IF sy-subrc = 0.
        CALL METHOD (ls_seoclass-clsname)=>xstring_to_solix
          EXPORTING
            iv_xstring = e_file
          RECEIVING
            et_solix   = et_file.
        e_bytecount = xstrlen( e_file ).
      ELSE.
        " Convert to binary
        CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
          EXPORTING
            buffer        = e_file
          IMPORTING
            output_length = e_bytecount
          TABLES
            binary_tab    = et_file.
      ENDIF.
    ENDIF.

  ENDMETHOD.