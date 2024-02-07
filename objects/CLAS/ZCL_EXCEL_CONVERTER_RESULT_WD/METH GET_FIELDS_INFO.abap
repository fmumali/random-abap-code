  METHOD get_fields_info.
    DATA: lo_aggr    TYPE REF TO cl_salv_wd_aggr_rule,
          l_aggrtype TYPE salv_wd_constant.

    FIELD-SYMBOLS: <fs_fields>  TYPE salv_wd_s_field_ref.

    READ TABLE wt_fields ASSIGNING <fs_fields> WITH KEY fieldname = xs_fcat-fieldname.
    IF sy-subrc = 0.
      lo_aggr = <fs_fields>-r_field->if_salv_wd_aggr~get_aggr_rule( ) .
      IF lo_aggr IS BOUND.
        l_aggrtype = lo_aggr->get_aggregation_type( ) .
        CASE l_aggrtype.
          WHEN if_salv_wd_c_aggregation=>aggrtype_total.
            xs_fcat-do_sum = abap_true.
          WHEN if_salv_wd_c_aggregation=>aggrtype_minimum.
            xs_fcat-do_sum =  'A'.
          WHEN if_salv_wd_c_aggregation=>aggrtype_maximum .
            xs_fcat-do_sum =  'B'.
          WHEN if_salv_wd_c_aggregation=>aggrtype_average .
            xs_fcat-do_sum =  'C'.
          WHEN OTHERS.
            CLEAR xs_fcat-do_sum .
        ENDCASE.
      ENDIF.
    ENDIF.

  ENDMETHOD.