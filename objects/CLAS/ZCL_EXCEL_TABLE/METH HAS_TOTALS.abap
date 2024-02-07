  METHOD has_totals.
    DATA: ls_field_catalog    TYPE zexcel_s_fieldcatalog.

    ep_result = abap_false.

    LOOP AT fieldcat INTO ls_field_catalog.
      IF ls_field_catalog-totals_function IS NOT INITIAL.
        ep_result = abap_true.
        EXIT.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.