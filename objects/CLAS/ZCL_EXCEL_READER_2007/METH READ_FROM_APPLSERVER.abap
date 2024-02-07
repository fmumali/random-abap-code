  METHOD read_from_applserver.

    DATA: lv_filelength         TYPE i,
          lt_binary_data        TYPE STANDARD TABLE OF x255 WITH NON-UNIQUE DEFAULT KEY,
          ls_binary_data        LIKE LINE OF lt_binary_data,
          lv_filename           TYPE string,
          lv_max_length_line    TYPE i,
          lv_actual_length_line TYPE i,
          lv_errormessage       TYPE string.

    lv_filename = i_filename.

    DESCRIBE FIELD ls_binary_data LENGTH lv_max_length_line IN BYTE MODE.
    OPEN DATASET lv_filename FOR INPUT IN BINARY MODE.
    IF sy-subrc <> 0.
      lv_errormessage = 'A problem occured when reading the file'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.
    WHILE sy-subrc = 0.

      READ DATASET lv_filename INTO ls_binary_data MAXIMUM LENGTH lv_max_length_line ACTUAL LENGTH lv_actual_length_line.
      APPEND ls_binary_data TO lt_binary_data.
      lv_filelength = lv_filelength + lv_actual_length_line.

    ENDWHILE.
    CLOSE DATASET lv_filename.

*--------------------------------------------------------------------*
* Binary data needs to be provided as XSTRING for further processing
*--------------------------------------------------------------------*
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_filelength
      IMPORTING
        buffer       = r_excel_data
      TABLES
        binary_tab   = lt_binary_data.
  ENDMETHOD.