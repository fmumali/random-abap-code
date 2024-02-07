  METHOD create_table.
    TYPES: BEGIN OF ts_output,
             fieldname TYPE fieldname,
             function  TYPE funcname,
           END OF ts_output.

    DATA: lo_data TYPE REF TO data.
    DATA: lo_addit          TYPE REF TO cl_abap_elemdescr,
          lt_components_tab TYPE cl_abap_structdescr=>component_table,
          ls_components     TYPE abap_componentdescr,
          lo_table          TYPE REF TO cl_abap_tabledescr,
          lo_struc          TYPE REF TO cl_abap_structdescr.

    FIELD-SYMBOLS: <fs_scat>  TYPE zexcel_s_converter_fcat,
                   <fs_stab>  TYPE any,
                   <fs_ttab>  TYPE STANDARD TABLE,
                   <fs>       TYPE any,
                   <fs_table> TYPE STANDARD TABLE.

    SORT wt_fieldcatalog BY position.
    ASSIGN wo_table->* TO <fs_table>.

    READ TABLE <fs_table> ASSIGNING <fs_stab> INDEX 1.
    IF sy-subrc EQ 0 .
      LOOP AT wt_fieldcatalog ASSIGNING <fs_scat>.
        ASSIGN COMPONENT <fs_scat>-columnname OF STRUCTURE <fs_stab> TO <fs>.
        IF sy-subrc = 0.
          ls_components-name   = <fs_scat>-columnname.
          TRY.
              lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
            CATCH cx_sy_move_cast_error.
              CLEAR lo_addit.
              DELETE TABLE wt_fieldcatalog FROM <fs_scat>.
          ENDTRY.
          IF lo_addit IS BOUND.
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
          ENDIF.
        ENDIF.
      ENDLOOP.
      IF lt_components_tab IS NOT INITIAL.
        "create new line type
        TRY.
            lo_struc = cl_abap_structdescr=>create( p_components = lt_components_tab
                                                    p_strict     = abap_false ).
          CATCH cx_sy_struct_creation.
            RETURN.  " We can not do anything in this case.
        ENDTRY.

        lo_table = cl_abap_tabledescr=>create( lo_struc ).

        CREATE DATA wo_data   TYPE HANDLE lo_table.
        CREATE DATA lo_data   TYPE HANDLE lo_struc.

        ASSIGN wo_data->* TO <fs_ttab>.
        ASSIGN lo_data->* TO <fs_stab>.
        LOOP AT <fs_table>  ASSIGNING <fs>.
          CLEAR <fs_stab>.
          MOVE-CORRESPONDING <fs> TO <fs_stab>.
          APPEND <fs_stab> TO <fs_ttab>.
        ENDLOOP.
      ENDIF.
    ENDIF.

  ENDMETHOD.