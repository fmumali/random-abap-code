  METHOD describe_structure.
    DATA: lt_components TYPE abap_component_tab,
          lt_comps      TYPE abap_component_tab,
          ls_component  TYPE abap_componentdescr,
          lo_elemdescr  TYPE REF TO cl_abap_elemdescr,
          ls_dfies      TYPE dfies,
          l_position    LIKE ls_dfies-position.

    "for DDIC structure get the info directly
    IF io_struct->is_ddic_type( ) = abap_true.
      rt_dfies = io_struct->get_ddic_field_list( ).
    ELSE.
      lt_components = io_struct->get_components( ).

      LOOP AT lt_components INTO ls_component.
        structure_case( EXPORTING is_component  = ls_component
                        CHANGING  xt_components = lt_comps   ) .
      ENDLOOP.
      LOOP AT lt_comps INTO ls_component.
        CLEAR ls_dfies.
        IF ls_component-type->kind = cl_abap_typedescr=>kind_elem. "E Elementary Type
          ADD 1 TO l_position.
          lo_elemdescr ?= ls_component-type.
          IF lo_elemdescr->is_ddic_type( ) = abap_true.
            ls_dfies           = lo_elemdescr->get_ddic_field( ).
            ls_dfies-fieldname = ls_component-name.
            ls_dfies-position  = l_position.
          ELSE.
            ls_dfies-fieldname = ls_component-name.
            ls_dfies-position  = l_position.
            ls_dfies-inttype   = lo_elemdescr->type_kind.
            ls_dfies-leng      = lo_elemdescr->length.
            ls_dfies-outputlen = lo_elemdescr->length.
            ls_dfies-decimals  = lo_elemdescr->decimals.
            ls_dfies-fieldtext = ls_component-name.
            ls_dfies-reptext   = ls_component-name.
            ls_dfies-scrtext_s = ls_component-name.
            ls_dfies-scrtext_m = ls_component-name.
            ls_dfies-scrtext_l = ls_component-name.
            ls_dfies-dynpfld   = abap_true.
          ENDIF.
          INSERT ls_dfies INTO TABLE rt_dfies.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.