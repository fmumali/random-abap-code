  PROTECTED SECTION.

    TYPES:
      BEGIN OF t_relationship,
        id     TYPE string,
        type   TYPE string,
        target TYPE string,
      END OF t_relationship .
    TYPES:
      BEGIN OF t_fileversion,
        appname      TYPE string,
        lastedited   TYPE string,
        lowestedited TYPE string,
        rupbuild     TYPE string,
        codename     TYPE string,
      END OF t_fileversion .
    TYPES:
      BEGIN OF t_sheet,
        name    TYPE string,
        sheetid TYPE string,
        id      TYPE string,
      END OF t_sheet .
    TYPES:
      BEGIN OF t_workbookpr,
        codename            TYPE string,
        defaultthemeversion TYPE string,
      END OF t_workbookpr .
    TYPES:
      BEGIN OF t_sheetpr,
        codename TYPE string,
      END OF t_sheetpr .

    DATA w_row_int TYPE zexcel_cell_row VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA w_col_int TYPE zexcel_cell_column VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
*"* private components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!