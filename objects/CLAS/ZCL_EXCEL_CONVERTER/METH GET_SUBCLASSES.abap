  METHOD get_subclasses.
    DATA:
      lt_subclasses TYPE seor_inheritance_keys,
      ls_subclass   TYPE seor_inheritance_key,
      lr_classdescr TYPE REF TO cl_abap_classdescr.

    CALL FUNCTION 'SEO_CLASS_GET_ALL_SUBS'
      EXPORTING
        clskey             = is_clskey
      IMPORTING
        inhkeys            = lt_subclasses
      EXCEPTIONS
        class_not_existing = 1
        OTHERS             = 2.

    CHECK sy-subrc = 0.

    LOOP AT lt_subclasses INTO ls_subclass.
      lr_classdescr ?= cl_abap_classdescr=>describe_by_name( ls_subclass-clsname ).
      IF lr_classdescr->is_instantiatable( ) = abap_true.
        APPEND ls_subclass TO ct_classes.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.