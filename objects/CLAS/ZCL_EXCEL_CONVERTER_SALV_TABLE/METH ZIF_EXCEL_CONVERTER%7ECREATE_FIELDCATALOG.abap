  METHOD zif_excel_converter~create_fieldcatalog.
    DATA: lo_salv TYPE REF TO cl_salv_table.

    zif_excel_converter~can_convert_object( io_object = io_object ).

    ws_option = is_option.

    lo_salv ?= io_object.

    CLEAR: es_layout,
           et_fieldcatalog,
           et_colors .

    IF lo_salv IS BOUND.
      load_data( EXPORTING io_salv   = lo_salv
                           it_table  = it_table ).
      apply_sort( EXPORTING it_table = it_table
                  IMPORTING eo_table = eo_table ) .

      get_color( EXPORTING io_table    = eo_table
                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    ENDIF.
  ENDMETHOD.