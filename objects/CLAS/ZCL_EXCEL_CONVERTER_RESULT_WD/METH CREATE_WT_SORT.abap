  METHOD create_wt_sort.
    DATA: lo_sort      TYPE REF TO cl_salv_wd_sort_rule,
          l_sort_order TYPE salv_wd_constant,
          ls_sort      TYPE lvc_s_sort.

    FIELD-SYMBOLS: <fs_fields>  TYPE salv_wd_s_field_ref.

    LOOP AT  wt_fields ASSIGNING <fs_fields>.
      lo_sort      = <fs_fields>-r_field->if_salv_wd_sort~get_sort_rule( ) .
      IF lo_sort IS BOUND.
        l_sort_order = lo_sort->get_sort_order( ).
        IF l_sort_order <> if_salv_wd_c_sort=>sort_order.
          CLEAR ls_sort.
          ls_sort-spos      = lo_sort->get_sort_position( ).
          ls_sort-fieldname = <fs_fields>-fieldname.
          ls_sort-subtot    = lo_sort->get_group_aggregation( ).
          IF l_sort_order = if_salv_wd_c_sort=>sort_order_ascending.
            ls_sort-up = abap_true.
          ELSE.
            ls_sort-down = abap_true.
          ENDIF.
          INSERT ls_sort INTO TABLE wt_sort.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.