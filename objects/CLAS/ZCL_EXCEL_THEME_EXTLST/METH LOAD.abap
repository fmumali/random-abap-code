  METHOD load.
    extlst = zcl_excel_common=>clone_ixml_with_namespaces( io_extlst ).
  ENDMETHOD.                    "load