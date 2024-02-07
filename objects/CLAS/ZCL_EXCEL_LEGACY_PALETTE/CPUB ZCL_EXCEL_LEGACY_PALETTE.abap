CLASS zcl_excel_legacy_palette DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_LEGACY_PALETTE
*"* do not include other source files here!!!

    METHODS constructor .
    METHODS is_modified
      RETURNING
        VALUE(ep_modified) TYPE abap_bool .
    METHODS get_color
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(ep_color) TYPE zexcel_style_color_argb
      RAISING
        zcx_excel .
    METHODS get_colors
      RETURNING
        VALUE(ep_colors) TYPE zexcel_t_style_color_argb .
    METHODS set_color
      IMPORTING
        !ip_index TYPE i
        !ip_color TYPE zexcel_style_color_argb
      RAISING
        zcx_excel .