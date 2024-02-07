  METHOD create_wt_fcat.
    DATA: lr_data        TYPE REF TO data,
          lo_structdescr TYPE REF TO cl_abap_structdescr,
          lt_dfies       TYPE ddfields,
          ls_dfies       TYPE dfies.

    DATA: ls_fcat         TYPE lvc_s_fcat.

    FIELD-SYMBOLS: <fs_tab>         TYPE ANY TABLE.

    ASSIGN io_table->* TO <fs_tab> .
    CREATE DATA lr_data LIKE LINE OF <fs_tab>.

    lo_structdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

    lt_dfies = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).

    LOOP AT lt_dfies INTO ls_dfies.
      MOVE-CORRESPONDING ls_dfies TO ls_fcat.
      ls_fcat-col_pos = ls_dfies-position.
      ls_fcat-key     = ls_dfies-keyflag.
      get_fields_info( CHANGING xs_fcat = ls_fcat ) .

      ls_fcat-col_opt = abap_true.

      get_columns_info( CHANGING xs_fcat = ls_fcat ) .

      INSERT ls_fcat INTO TABLE wt_fcat.
    ENDLOOP.

  ENDMETHOD.