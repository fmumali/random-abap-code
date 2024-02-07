  METHOD zif_excel_converter~create_fieldcatalog.
    DATA: lo_alv TYPE REF TO cl_gui_alv_grid.

    zif_excel_converter~can_convert_object( io_object = io_object ).

    ws_option = is_option.

    lo_alv ?= io_object.

    CLEAR: es_layout,
           et_fieldcatalog.

    IF lo_alv IS BOUND.
      lo_alv->get_frontend_fieldcatalog( IMPORTING et_fieldcatalog = wt_fcat ).
      lo_alv->get_frontend_layout( IMPORTING es_layout = ws_layo ).
      lo_alv->get_sort_criteria( IMPORTING et_sort = wt_sort ) .
      lo_alv->get_filter_criteria( IMPORTING et_filter = wt_filt ) .

      apply_sort( EXPORTING it_table = it_table
                  IMPORTING eo_table = eo_table ) .

      get_color( EXPORTING io_table    = eo_table
                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table  ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    ENDIF.
  ENDMETHOD.