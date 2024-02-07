  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_READER_XLSM
*"* do not include other source files here!!!
    METHODS load_vbaproject
      IMPORTING
        !ip_path  TYPE string
        !ip_excel TYPE REF TO zcl_excel
      RAISING
        zcx_excel .