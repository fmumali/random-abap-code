  METHOD load.
    objectdefaults = zcl_excel_common=>clone_ixml_with_namespaces( io_object_def ).
  ENDMETHOD.                    "load