  METHOD load.
    fmt_scheme = zcl_excel_common=>clone_ixml_with_namespaces( io_fmt_scheme ).
  ENDMETHOD.                    "load