  METHOD load_data.
    DATA: lo_columns      TYPE REF TO cl_salv_columns_table,
          lo_aggregations TYPE REF TO cl_salv_aggregations,
          lo_sorts        TYPE REF TO cl_salv_sorts,
          lo_filters      TYPE REF TO cl_salv_filters,
          lo_functional   TYPE REF TO cl_salv_functional_settings,
          lo_display      TYPE REF TO cl_salv_display_settings.

    DATA: ls_vari   TYPE disvariant,
          lo_layout TYPE REF TO cl_salv_layout.

    DATA lt_kkblo_fieldcat TYPE kkblo_t_fieldcat.
    DATA ls_kkblo_layout   TYPE kkblo_layout.
    DATA lt_kkblo_filter   TYPE kkblo_t_filter.
    DATA lt_kkblo_sort     TYPE kkblo_t_sortinfo.
    DATA: lv_intercept_data_active TYPE abap_bool,
          ls_layout_key            TYPE salv_s_layout_key.

    lo_layout               = io_salv->get_layout( ) .
    lo_columns              = io_salv->get_columns( ).
    lo_aggregations         = io_salv->get_aggregations( ) .
    lo_sorts                = io_salv->get_sorts( ) .
    lo_filters              = io_salv->get_filters( ) .
    lo_display              = io_salv->get_display_settings( ) .
    lo_functional           = io_salv->get_functional_settings( ) .

    CLEAR: wt_fcat, wt_sort, wt_filt.

    lv_intercept_data_active = is_intercept_data_active( ).

* First update metadata if we can.
    IF io_salv->is_offline( ) = abap_false.
      IF lv_intercept_data_active = abap_true.
        ls_layout_key = lo_layout->get_key( ).
        ls_vari-report    = ls_layout_key-report.
        ls_vari-handle    = ls_layout_key-handle.
        ls_vari-log_group = ls_layout_key-logical_group.
        ls_vari-variant   = lo_layout->get_initial_layout( ).
      ELSE.
        IF zcl_excel_converter_salv_model=>is_get_metadata_callable( io_salv ) = abap_true.
          io_salv->get_metadata( ) .
        ELSE.
          " (do same as offline below)
          cl_salv_controller_metadata=>get_variant(
            EXPORTING
              r_layout  = lo_layout
            CHANGING
              s_variant = ls_vari ).
        ENDIF.
      ENDIF.
    ELSE.
* If we are offline we need to build this.
      cl_salv_controller_metadata=>get_variant(
        EXPORTING
          r_layout  = lo_layout
        CHANGING
          s_variant = ls_vari ).
    ENDIF.

*... get the column information
    wt_fcat = cl_salv_controller_metadata=>get_lvc_fieldcatalog(
                           r_columns      = lo_columns
                           r_aggregations = lo_aggregations ).

*... get the layout information
    cl_salv_controller_metadata=>get_lvc_layout(
      EXPORTING
        r_functional_settings = lo_functional
        r_display_settings    = lo_display
        r_columns             = lo_columns
        r_aggregations        = lo_aggregations
      CHANGING
        s_layout              = ws_layo ).

* the fieldcatalog is not complete yet!
    CALL FUNCTION 'LVC_FIELDCAT_COMPLETE'
      EXPORTING
        i_complete       = 'X'
        i_refresh_buffer = space
        i_buffer_active  = space
        is_layout        = ws_layo
        i_test           = '1'
        i_fcat_complete  = 'X'
      IMPORTING
        es_layout        = ws_layo
      CHANGING
        ct_fieldcat      = wt_fcat.

    IF ls_vari IS NOT INITIAL AND
        ( io_salv->is_offline( ) = abap_true
          OR lv_intercept_data_active = abap_true ).
      CALL FUNCTION 'LVC_TRANSFER_TO_KKBLO'
        EXPORTING
          it_fieldcat_lvc         = wt_fcat
          is_layout_lvc           = ws_layo
        IMPORTING
          et_fieldcat_kkblo       = lt_kkblo_fieldcat
          es_layout_kkblo         = ls_kkblo_layout
        TABLES
          it_data                 = it_table
        EXCEPTIONS
          it_data_missing         = 1
          it_fieldcat_lvc_missing = 2
          OTHERS                  = 3.
      IF sy-subrc <> 0.
      ENDIF.

      CALL FUNCTION 'LT_VARIANT_LOAD'
        EXPORTING
          i_tabname           = '1'
          i_dialog            = ' '
          i_user_specific     = 'X'
          i_fcat_complete     = 'X'
        IMPORTING
          et_fieldcat         = lt_kkblo_fieldcat
          et_sort             = lt_kkblo_sort
          et_filter           = lt_kkblo_filter
        CHANGING
          cs_layout           = ls_kkblo_layout
          ct_default_fieldcat = lt_kkblo_fieldcat
          cs_variant          = ls_vari
        EXCEPTIONS
          wrong_input         = 1
          fc_not_complete     = 2
          not_found           = 3
          OTHERS              = 4.
      IF sy-subrc <> 0.
      ENDIF.

      CALL FUNCTION 'LVC_TRANSFER_FROM_KKBLO'
        EXPORTING
          it_fieldcat_kkblo = lt_kkblo_fieldcat
          it_sort_kkblo     = lt_kkblo_sort
          it_filter_kkblo   = lt_kkblo_filter
          is_layout_kkblo   = ls_kkblo_layout
        IMPORTING
          et_fieldcat_lvc   = wt_fcat
          et_sort_lvc       = wt_sort
          et_filter_lvc     = wt_filt
          es_layout_lvc     = ws_layo
        TABLES
          it_data           = it_table
        EXCEPTIONS
          it_data_missing   = 1
          OTHERS            = 2.
      IF sy-subrc <> 0.
      ENDIF.

    ELSE.
*  ... get the sort information
      wt_sort = cl_salv_controller_metadata=>get_lvc_sort( lo_sorts ).

*  ... get the filter information
      wt_filt = cl_salv_controller_metadata=>get_lvc_filter( lo_filters ).
    ENDIF.

  ENDMETHOD.