  METHOD create_xl_theme.
    DATA: lo_theme TYPE REF TO zcl_excel_theme.

    excel->get_theme(
    IMPORTING
      eo_theme = lo_theme
      ).
    IF lo_theme IS INITIAL.
      CREATE OBJECT lo_theme.
    ENDIF.
    ep_content = lo_theme->write_theme( ).

  ENDMETHOD.