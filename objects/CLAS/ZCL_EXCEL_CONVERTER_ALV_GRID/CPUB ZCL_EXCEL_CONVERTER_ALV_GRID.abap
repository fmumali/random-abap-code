CLASS zcl_excel_converter_alv_grid DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_converter_alv
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS zif_excel_converter~can_convert_object
        REDEFINITION .
*"* public components of class ZCL_EXCEL_CONVERTER_ALV_GRID
*"* do not include other source files here!!!
    METHODS zif_excel_converter~create_fieldcatalog
        REDEFINITION .
    METHODS zif_excel_converter~get_supported_class
        REDEFINITION .