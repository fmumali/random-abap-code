  METHOD get_columns_info.
    DATA:  l_numc2             TYPE salv_wd_constant.


    FIELD-SYMBOLS: <fs_column>  TYPE salv_wd_s_column_ref.

    READ TABLE wt_columns ASSIGNING <fs_column> WITH KEY id = xs_fcat-fieldname .
    IF sy-subrc = 0.
      xs_fcat-col_pos    = <fs_column>-r_column->get_position( ) .
      l_numc2 = <fs_column>-r_column->get_fixed_position( ).
      IF l_numc2 = '02'.
        xs_fcat-fix_column = abap_true .
      ENDIF.
      l_numc2 = <fs_column>-r_column->get_visible( ).
      IF l_numc2 = '01'.
        xs_fcat-no_out = abap_true .
      ENDIF.
    ENDIF.

  ENDMETHOD.