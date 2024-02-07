  METHOD create_xl_drawings_vml.

    DATA:
      ld_stream       TYPE string.


* INIT_RESULT
    CLEAR ep_content.


* BODY
    ld_stream = set_vml_string( ).

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = ld_stream
      IMPORTING
        buffer = ep_content
      EXCEPTIONS
        failed = 1
        OTHERS = 2.
    IF sy-subrc <> 0.
      CLEAR ep_content.
    ENDIF.


  ENDMETHOD.