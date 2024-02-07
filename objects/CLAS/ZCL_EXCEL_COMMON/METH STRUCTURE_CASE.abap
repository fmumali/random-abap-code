  METHOD structure_case.
    DATA: lt_comp_str        TYPE abap_component_tab.

    CASE is_component-type->kind.
      WHEN cl_abap_typedescr=>kind_elem. "E Elementary Type
        INSERT is_component INTO TABLE xt_components.
      WHEN cl_abap_typedescr=>kind_table. "T Table
        INSERT is_component INTO TABLE xt_components.
      WHEN cl_abap_typedescr=>kind_struct. "S Structure
        lt_comp_str = structure_recursive( is_component = is_component ).
        INSERT LINES OF lt_comp_str INTO TABLE xt_components.
      WHEN OTHERS. "cl_abap_typedescr=>kind_ref or  cl_abap_typedescr=>kind_class or  cl_abap_typedescr=>kind_intf.
* We skip it. for now.
    ENDCASE.
  ENDMETHOD.