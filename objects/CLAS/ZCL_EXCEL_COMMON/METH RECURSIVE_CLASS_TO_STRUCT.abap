  METHOD recursive_class_to_struct.
    " # issue 139
* is working for me - but after looking through this coding I guess
* I'll rewrite this to a version w/o recursion
* This is private an no one using it so far except me, so no need to hurry
    DATA: descr          TYPE REF TO cl_abap_structdescr,
          wa_component   LIKE LINE OF descr->components,
          attribute_name LIKE wa_component-name,
          flag_class     TYPE abap_bool.

    FIELD-SYMBOLS: <field>     TYPE any,
                   <fieldx>    TYPE any,
                   <attribute> TYPE any.


    descr ?= cl_abap_structdescr=>describe_by_data( e_target ).

    LOOP AT descr->components INTO wa_component.

* Assign structure and X-structure
      ASSIGN COMPONENT wa_component-name OF STRUCTURE e_target  TO <field>.
      ASSIGN COMPONENT wa_component-name OF STRUCTURE e_targetx TO <fieldx>.
* At least one field in the structure should be marked - otherwise continue with next field
      CLEAR flag_class.
* maybe source is just a structure - try assign component...
      ASSIGN COMPONENT wa_component-name OF STRUCTURE i_source  TO <attribute>.
      IF sy-subrc <> 0.
* not - then it is an attribute of the class - use different assign then
        CONCATENATE 'i_source->' wa_component-name INTO attribute_name.
        ASSIGN (attribute_name) TO <attribute>.
        IF sy-subrc <> 0.
          EXIT.
        ENDIF.  " Should not happen if structure is built properly - otherwise just exit to create no dumps
        flag_class = abap_true.
      ENDIF.

      CASE wa_component-type_kind.
        WHEN cl_abap_structdescr=>typekind_struct1 OR cl_abap_structdescr=>typekind_struct2.  " Structure --> use recursio
          zcl_excel_common=>recursive_class_to_struct( EXPORTING i_source  = <attribute>
                                                       CHANGING  e_target  = <field>
                                                                 e_targetx = <fieldx> ).
        WHEN OTHERS.
          <field> = <attribute>.
          <fieldx> = abap_true.

      ENDCASE.
    ENDLOOP.

  ENDMETHOD.