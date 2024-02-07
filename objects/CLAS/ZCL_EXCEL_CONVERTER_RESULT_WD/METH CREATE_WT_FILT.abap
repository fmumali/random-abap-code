  METHOD create_wt_filt.
* No neeed for superclass.
* Only for WD
    DATA: lt_filters TYPE salv_wd_t_filter_rule_ref,
          ls_filt    TYPE lvc_s_filt.

    FIELD-SYMBOLS: <fs_fields> TYPE salv_wd_s_field_ref,
                   <fs_filter> TYPE salv_wd_s_filter_rule_ref.

    LOOP AT  wt_fields ASSIGNING <fs_fields>.
      CLEAR lt_filters.
      lt_filters    = <fs_fields>-r_field->if_salv_wd_filter~get_filter_rules( ) .
      LOOP AT lt_filters ASSIGNING <fs_filter>.
        ls_filt-fieldname = <fs_fields>-fieldname.
        IF <fs_filter>-r_filter_rule->get_included( ) = abap_true.
          ls_filt-sign      = 'I'.
        ELSE.
          ls_filt-sign      = 'E'.
        ENDIF.
        ls_filt-option    = <fs_filter>-r_filter_rule->get_operator( ).
        ls_filt-high      = <fs_filter>-r_filter_rule->get_high_value( ) .
        ls_filt-low       = <fs_filter>-r_filter_rule->get_low_value( ) .
        INSERT ls_filt INTO TABLE wt_filt.
      ENDLOOP.
    ENDLOOP.

  ENDMETHOD.