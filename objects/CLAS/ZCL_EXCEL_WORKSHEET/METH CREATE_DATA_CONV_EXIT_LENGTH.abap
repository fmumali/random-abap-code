  METHOD create_data_conv_exit_length.
    DATA: lo_addit    TYPE REF TO cl_abap_elemdescr,
          ls_dfies    TYPE dfies,
          l_function  TYPE funcname,
          l_value(50) TYPE c.

    lo_addit ?= cl_abap_typedescr=>describe_by_data( ip_value ).
    lo_addit->get_ddic_field( RECEIVING  p_flddescr   = ls_dfies
                              EXCEPTIONS not_found    = 1
                                         no_ddic_type = 2
                                         OTHERS       = 3 ) .
    IF sy-subrc = 0 AND ls_dfies-convexit IS NOT INITIAL.
      CREATE DATA ep_value TYPE c LENGTH ls_dfies-outputlen.
    ELSE.
      CREATE DATA ep_value LIKE ip_value.
    ENDIF.

  ENDMETHOD.