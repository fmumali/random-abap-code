  METHOD recursive_struct_to_class.
    " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
    DATA: descr          TYPE REF TO cl_abap_structdescr,
          wa_component   LIKE LINE OF descr->components,
          attribute_name LIKE wa_component-name,
          flag_class     TYPE abap_bool,
          o_border       TYPE REF TO zcl_excel_style_border.

    FIELD-SYMBOLS: <field>     TYPE any,
                   <fieldx>    TYPE any,
                   <attribute> TYPE any.


    descr ?= cl_abap_structdescr=>describe_by_data( i_source ).

    LOOP AT descr->components INTO wa_component.

* Assign structure and X-structure
      ASSIGN COMPONENT wa_component-name OF STRUCTURE i_source  TO <field>.
      ASSIGN COMPONENT wa_component-name OF STRUCTURE i_sourcex TO <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
      CHECK <fieldx> CA abap_true.
      CLEAR flag_class.
* maybe target is just a structure - try assign component...
      ASSIGN COMPONENT wa_component-name OF STRUCTURE e_target  TO <attribute>.
      IF sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
        CONCATENATE 'E_TARGET->' wa_component-name INTO attribute_name.
        ASSIGN (attribute_name) TO <attribute>.
        IF sy-subrc <> 0.EXIT.ENDIF.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
        flag_class = abap_true.
      ENDIF.

      CASE wa_component-type_kind.
        WHEN cl_abap_structdescr=>typekind_struct1 OR cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursion
          " To avoid dump with attribute GRADTYPE of class ZCL_EXCEL_STYLE_FILL
          " quick and really dirty fix -> check the attribute name
          " Border has to be initialized somewhere else
          IF wa_component-name EQ 'GRADTYPE'.
            flag_class = abap_false.
          ENDIF.

          IF flag_class = abap_true AND <attribute> IS INITIAL.
* Only borders will be passed as unbound references.  But since we want to set a value we have to create an instance
            CREATE OBJECT o_border.
            <attribute> = o_border.
          ENDIF.
          zcl_excel_common=>recursive_struct_to_class( EXPORTING i_source  = <field>
                                                                 i_sourcex = <fieldx>
                                                       CHANGING  e_target  = <attribute> ).
        WHEN OTHERS.
          CHECK <fieldx> = abap_true.  " Marked for change
          <attribute> = <field>.

      ENDCASE.
    ENDLOOP.

  ENDMETHOD.