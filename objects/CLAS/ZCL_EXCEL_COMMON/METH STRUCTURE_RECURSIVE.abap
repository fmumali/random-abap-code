  METHOD structure_recursive.
    DATA: lo_struct     TYPE REF TO cl_abap_structdescr,
          lt_components TYPE abap_component_tab,
          ls_components TYPE abap_componentdescr.

    lo_struct ?= is_component-type.
    lt_components = lo_struct->get_components( ).

    LOOP AT lt_components INTO ls_components.
      structure_case( EXPORTING is_component  = ls_components
                      CHANGING  xt_components = rt_components ) .
    ENDLOOP.

  ENDMETHOD.