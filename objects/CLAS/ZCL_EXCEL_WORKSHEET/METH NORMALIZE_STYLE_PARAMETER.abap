  METHOD normalize_style_parameter.

    DATA: lo_style_type TYPE REF TO cl_abap_typedescr.
    FIELD-SYMBOLS:
      <style> TYPE REF TO zcl_excel_style.

    CHECK ip_style_or_guid IS NOT INITIAL.

    lo_style_type = cl_abap_typedescr=>describe_by_data( ip_style_or_guid ).
    IF lo_style_type->type_kind = lo_style_type->typekind_oref.
      lo_style_type = cl_abap_typedescr=>describe_by_object_ref( ip_style_or_guid ).
      IF lo_style_type->absolute_name = '\CLASS=ZCL_EXCEL_STYLE'.
        ASSIGN ip_style_or_guid TO <style>.
        rv_guid = <style>->get_guid( ).
      ENDIF.

    ELSEIF lo_style_type->absolute_name = '\TYPE=ZEXCEL_CELL_STYLE'.
      rv_guid = ip_style_or_guid.

    ELSEIF lo_style_type->type_kind = lo_style_type->typekind_hex.
      rv_guid = ip_style_or_guid.

    ELSE.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'IP_GUID type must be either REF TO zcl_excel_style or zexcel_cell_style'.
    ENDIF.

  ENDMETHOD.