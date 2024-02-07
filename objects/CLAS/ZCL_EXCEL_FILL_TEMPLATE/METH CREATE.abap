  METHOD create.

    CREATE OBJECT eo_template_filler .

    eo_template_filler->mo_excel = io_excel.
    eo_template_filler->get_range( ).
    eo_template_filler->discard_overlapped( ).
    eo_template_filler->sign_range( ).
    eo_template_filler->find_var( ).

  ENDMETHOD.