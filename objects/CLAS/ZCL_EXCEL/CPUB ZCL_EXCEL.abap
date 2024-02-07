CLASS zcl_excel DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL
*"* do not include other source files here!!!
    INTERFACES zif_excel_book_properties .
    INTERFACES zif_excel_book_protection .
    INTERFACES zif_excel_book_vba_project .

    DATA legacy_palette TYPE REF TO zcl_excel_legacy_palette READ-ONLY .
    DATA security TYPE REF TO zcl_excel_security .
    DATA use_template TYPE abap_bool .
    CONSTANTS version TYPE c LENGTH 10 VALUE '7.16.0'.      "#EC NOTEXT

    METHODS add_new_autofilter
      IMPORTING
        !io_sheet            TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter
      RAISING
        zcx_excel .
    METHODS add_new_comment
      RETURNING
        VALUE(eo_comment) TYPE REF TO zcl_excel_comment .
    METHODS add_new_drawing
      IMPORTING
        !ip_type          TYPE zexcel_drawing_type DEFAULT zcl_excel_drawing=>type_image
        !ip_title         TYPE clike OPTIONAL
      RETURNING
        VALUE(eo_drawing) TYPE REF TO zcl_excel_drawing .
    METHODS add_new_range
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS add_new_style
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style OPTIONAL
        !io_clone_of    TYPE REF TO zcl_excel_style OPTIONAL
          PREFERRED PARAMETER !ip_guid
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS add_new_worksheet
      IMPORTING
        !ip_title           TYPE zexcel_sheet_title OPTIONAL
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS add_static_styles .
    METHODS constructor .
    METHODS delete_worksheet
      IMPORTING
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS delete_worksheet_by_index
      IMPORTING
        !iv_index TYPE numeric
      RAISING
        zcx_excel .
    METHODS delete_worksheet_by_name
      IMPORTING
        !iv_title TYPE clike
      RAISING
        zcx_excel .
    METHODS get_active_sheet_index
      RETURNING
        VALUE(r_active_worksheet) TYPE zexcel_active_worksheet .
    METHODS get_active_worksheet
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_autofilters_reference
      RETURNING
        VALUE(ro_autofilters) TYPE REF TO zcl_excel_autofilters .
    METHODS get_default_style
      RETURNING
        VALUE(ep_style) TYPE zexcel_cell_style .
    METHODS get_drawings_iterator
      IMPORTING
        !ip_type           TYPE zexcel_drawing_type
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_next_table_id
      RETURNING
        VALUE(ep_id) TYPE i .
    METHODS get_ranges_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_static_cellstyle_guid
      IMPORTING
        !ip_cstyle_complete  TYPE zexcel_s_cstyle_complete
        !ip_cstylex_complete TYPE zexcel_s_cstylex_complete
      RETURNING
        VALUE(ep_guid)       TYPE zexcel_cell_style .
    METHODS get_styles_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_style_index_in_styles
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style
      RETURNING
        VALUE(ep_index) TYPE i
      RAISING
        zcx_excel .
    METHODS get_style_to_guid
      IMPORTING
        !ip_guid               TYPE zexcel_cell_style
      RETURNING
        VALUE(ep_stylemapping) TYPE zexcel_s_stylemapping
      RAISING
        zcx_excel .
    METHODS get_theme
      EXPORTING
        !eo_theme TYPE REF TO zcl_excel_theme .
    METHODS get_worksheets_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_worksheets_name
      RETURNING
        VALUE(ep_name) TYPE zexcel_worksheets_name .
    METHODS get_worksheets_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_worksheet_by_index
      IMPORTING
        !iv_index           TYPE numeric
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_worksheet_by_name
      IMPORTING
        !ip_sheet_name      TYPE zexcel_sheet_title
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS set_active_sheet_index
      IMPORTING
        !i_active_worksheet TYPE zexcel_active_worksheet
      RAISING
        zcx_excel .
    METHODS set_active_sheet_index_by_name
      IMPORTING
        !i_worksheet_name TYPE zexcel_worksheets_name .
    METHODS set_default_style
      IMPORTING
        !ip_style TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS set_theme
      IMPORTING
        !io_theme TYPE REF TO zcl_excel_theme .
    METHODS fill_template
      IMPORTING
        !iv_data TYPE REF TO zcl_excel_template_data
      RAISING
        zcx_excel .