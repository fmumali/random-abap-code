  METHOD load.
    extracolor = zcl_excel_common=>clone_ixml_with_namespaces( io_extra_color ).
  ENDMETHOD.                    "load