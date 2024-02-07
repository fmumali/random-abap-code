  METHOD zif_excel_converter~create_fieldcatalog.
    DATA: lo_result TYPE REF TO cl_salv_wd_result_data_table,
          lo_data   TYPE REF TO data.

    FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE.

    zif_excel_converter~can_convert_object( io_object = io_object ).

    ws_option = is_option.

    lo_result ?= io_object.

    CLEAR: es_layout,
           et_fieldcatalog.

    IF lo_result IS BOUND.
      lo_data = get_table( io_object = lo_result->r_model->r_data ).
      IF lo_data IS BOUND.
        ASSIGN lo_data->* TO <fs_table> .

        wo_config ?= lo_result->r_model->r_model.

        IF wo_config IS BOUND.
          wt_fields  = wo_config->if_salv_wd_field_settings~get_fields( ) .
          wt_columns = wo_config->if_salv_wd_column_settings~get_columns( ) .
        ENDIF.

        create_wt_fcat( io_table = lo_data ).
        create_wt_sort( ).
        create_wt_filt( ).

        apply_sort( EXPORTING it_table = <fs_table>
                    IMPORTING eo_table = eo_table ) .

        get_filter( IMPORTING et_filter  = et_filter
                    CHANGING  xo_table   = eo_table ) .

        update_catalog( CHANGING  cs_layout       = es_layout
                                  ct_fieldcatalog = et_fieldcatalog ).
      ELSE.
* We have a problem and should stop here
      ENDIF.
    ENDIF.
  ENDMETHOD.