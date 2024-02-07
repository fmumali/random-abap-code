  METHOD get_value_type.
    DATA: lo_addit    TYPE REF TO cl_abap_elemdescr,
          ls_dfies    TYPE dfies,
          l_function  TYPE funcname,
          l_value(50) TYPE c.

    ep_value = ip_value.
    ep_value_type = cl_abap_typedescr=>typekind_string. " Thats our default if something goes wrong.

    TRY.
        lo_addit            ?= cl_abap_typedescr=>describe_by_data( ip_value ).
      CATCH cx_sy_move_cast_error.
        CLEAR lo_addit.
    ENDTRY.
    IF lo_addit IS BOUND.
      lo_addit->get_ddic_field( RECEIVING  p_flddescr   = ls_dfies
                                EXCEPTIONS not_found    = 1
                                           no_ddic_type = 2
                                           OTHERS       = 3 ) .
      IF sy-subrc = 0.
        ep_value_type = ls_dfies-inttype.

        IF ls_dfies-convexit IS NOT INITIAL.
* We need to convert with output conversion function
          CONCATENATE 'CONVERSION_EXIT_' ls_dfies-convexit '_OUTPUT' INTO l_function.
          SELECT SINGLE funcname INTO l_function
                FROM tfdir
                WHERE funcname = l_function.
          IF sy-subrc = 0.
            CALL FUNCTION l_function
              EXPORTING
                input  = ip_value
              IMPORTING
                output = l_value
              EXCEPTIONS
                OTHERS = 1.
            IF sy-subrc <> 0.
            ELSE.
              TRY.
                  ep_value = l_value.
                CATCH cx_root.
                  ep_value = ip_value.
              ENDTRY.
            ENDIF.
          ENDIF.
        ENDIF.
      ELSE.
        ep_value_type = lo_addit->get_data_type_kind( ip_value ).
      ENDIF.
    ENDIF.

  ENDMETHOD.                    "GET_VALUE_TYPE