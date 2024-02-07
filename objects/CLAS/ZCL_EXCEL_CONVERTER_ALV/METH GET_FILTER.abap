  METHOD get_filter.
    DATA: ls_filt   TYPE lvc_s_filt,
          l_line    TYPE i,
          ls_filter TYPE zexcel_s_converter_fil.
    DATA: lo_addit          TYPE REF TO cl_abap_elemdescr,
          lt_components_tab TYPE cl_abap_structdescr=>component_table,
          ls_components     TYPE abap_componentdescr,
          lo_table          TYPE REF TO cl_abap_tabledescr,
          lo_struc          TYPE REF TO cl_abap_structdescr,
          lo_trange         TYPE REF TO data,
          lo_srange         TYPE REF TO data,
          lo_ltabdata       TYPE REF TO data.

    FIELD-SYMBOLS: <fs_tab>    TYPE STANDARD TABLE,
                   <fs_ltab>   TYPE STANDARD TABLE,
                   <fs_stab>   TYPE any,
                   <fs>        TYPE any,
                   <fs1>       TYPE any,
                   <fs_srange> TYPE any,
                   <fs_trange> TYPE STANDARD TABLE.

    IF ws_option-filter = abap_false.
      CLEAR et_filter.
      RETURN.
    ENDIF.

    ASSIGN xo_table->* TO <fs_tab>.

    CREATE DATA lo_ltabdata LIKE <fs_tab>.
    ASSIGN lo_ltabdata->* TO <fs_ltab>.

    LOOP AT wt_filt INTO ls_filt.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_line = sy-tabix.
        ASSIGN COMPONENT ls_filt-fieldname OF STRUCTURE <fs_stab> TO <fs>.
        IF sy-subrc = 0.
          IF l_line = 1.
            CLEAR lt_components_tab.
            ls_components-name   = 'SIGN'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( ls_filt-sign ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'OPTION'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( ls_filt-option ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'LOW'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'HIGH'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            "create new line type
            TRY.
                lo_struc = cl_abap_structdescr=>create( p_components = lt_components_tab
                                                        p_strict     = abap_false ).
              CATCH cx_sy_struct_creation.
                CONTINUE.
            ENDTRY.
            lo_table = cl_abap_tabledescr=>create( lo_struc ).

            CREATE DATA lo_trange  TYPE HANDLE lo_table.
            CREATE DATA lo_srange  TYPE HANDLE lo_struc.

            ASSIGN lo_trange->* TO <fs_trange>.
            ASSIGN lo_srange->* TO <fs_srange>.
          ENDIF.
          CLEAR <fs_trange>.
          ASSIGN COMPONENT 'SIGN'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-sign.
          ASSIGN COMPONENT 'OPTION' OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-option.
          ASSIGN COMPONENT 'LOW'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-low.
          ASSIGN COMPONENT 'HIGH'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-high.
          INSERT <fs_srange> INTO TABLE <fs_trange>.
          IF <fs> IN <fs_trange>.
            IF ws_option-filter = abap_true.
              ls_filter-rownumber   = l_line.
              ls_filter-columnname  = ls_filt-fieldname.
              INSERT ls_filter INTO TABLE et_filter.
            ELSE.
              INSERT <fs_stab> INTO TABLE <fs_ltab>.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDLOOP.
      IF ws_option-filter = abap_undefined.
        <fs_tab> = <fs_ltab>.
        CLEAR <fs_ltab>.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.