  METHOD get_alv_converters.
    DATA:
      lt_direct_implementations TYPE seor_implementing_keys,
      lt_all_implementations    TYPE seor_implementing_keys,
      ls_impkey                 TYPE seor_implementing_key,
      ls_classkey               TYPE seoclskey,
      lr_implementation         TYPE REF TO zif_excel_converter,
      ls_object                 TYPE ts_alv_types,
      lr_classdescr             TYPE REF TO cl_abap_classdescr.

    ls_classkey-clsname = 'ZIF_EXCEL_CONVERTER'.

    CALL FUNCTION 'SEO_INTERFACE_IMPLEM_GET_ALL'
      EXPORTING
        intkey  = ls_classkey
      IMPORTING
        impkeys = lt_direct_implementations
      EXCEPTIONS
        OTHERS  = 2.

    CHECK sy-subrc = 0.

    LOOP AT lt_direct_implementations INTO ls_impkey.
      lr_classdescr ?= cl_abap_classdescr=>describe_by_name( ls_impkey-clsname ).
      IF lr_classdescr->is_instantiatable( ) = abap_true.
        APPEND ls_impkey TO lt_all_implementations.
      ENDIF.

      ls_classkey-clsname = ls_impkey-clsname.
      get_subclasses( EXPORTING is_clskey = ls_classkey CHANGING ct_classes = lt_all_implementations ).
    ENDLOOP.

    SORT lt_all_implementations BY clsname.
    DELETE ADJACENT DUPLICATES FROM lt_all_implementations COMPARING clsname.

    LOOP AT lt_all_implementations INTO ls_impkey.
      CLEAR ls_object.
      CREATE OBJECT lr_implementation TYPE (ls_impkey-clsname).
      ls_object-seoclass = lr_implementation->get_supported_class( ).

      IF ls_object-seoclass IS NOT INITIAL.
        ls_object-clsname  = ls_impkey-clsname.
        INSERT ls_object INTO TABLE wt_objects.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.