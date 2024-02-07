CLASS zcl_excel_worksheet DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!

    INTERFACES zif_excel_sheet_printsettings .
    INTERFACES zif_excel_sheet_properties .
    INTERFACES zif_excel_sheet_protection .
    INTERFACES zif_excel_sheet_vba_project .

    TYPES:
      BEGIN OF  mty_s_outline_row,
        row_from  TYPE i,
        row_to    TYPE i,
        collapsed TYPE abap_bool,
      END OF mty_s_outline_row .
    TYPES:
      mty_ts_outlines_row TYPE SORTED TABLE OF mty_s_outline_row WITH UNIQUE KEY row_from row_to .
    TYPES:
      BEGIN OF mty_s_ignored_errors,
        "! Cell reference (e.g. "A1") or list like "A1 A2" or range "A1:G1"
        cell_coords           TYPE zexcel_cell_coords,
        "! Ignore errors when cells contain formulas that result in an error.
        eval_error            TYPE abap_bool,
        "! Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        two_digit_text_year   TYPE abap_bool,
        "! Ignore errors when numbers are formatted as text or are preceded by an apostrophe.
        number_stored_as_text TYPE abap_bool,
        "! Ignore errors when a formula in a region of your WorkSheet differs from other formulas in the same region.
        formula               TYPE abap_bool,
        "! Ignore errors when formulas omit certain cells in a region.
        formula_range         TYPE abap_bool,
        "! Ignore errors when unlocked cells contain formulas.
        unlocked_formula      TYPE abap_bool,
        "! Ignore errors when formulas refer to empty cells.
        empty_cell_reference  TYPE abap_bool,
        "! Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified.
        list_data_validation  TYPE abap_bool,
        "! Ignore errors when cells contain a value different from a calculated column formula.
        "! In other words, for a calculated column, a cell in that column is considered to have an error
        "! if its formula is different from the calculated column formula, or doesn't contain a formula at all.
        calculated_column     TYPE abap_bool,
      END OF mty_s_ignored_errors .
    TYPES:
      mty_th_ignored_errors TYPE HASHED TABLE OF mty_s_ignored_errors WITH UNIQUE KEY cell_coords .
    TYPES:
      BEGIN OF mty_s_column_formula,
        id                     TYPE i,
        column                 TYPE zexcel_cell_column,
        formula                TYPE string,
        table_top_left_row     TYPE zexcel_cell_row,
        table_bottom_right_row TYPE zexcel_cell_row,
        table_left_column_int  TYPE zexcel_cell_column,
        table_right_column_int TYPE zexcel_cell_column,
      END OF mty_s_column_formula .
    TYPES:
      mty_th_column_formula
               TYPE HASHED TABLE OF mty_s_column_formula
               WITH UNIQUE KEY id .
    TYPES:
      ty_doc_url TYPE c LENGTH 255 .
    TYPES:
      BEGIN OF mty_merge,
        row_from TYPE i,
        row_to   TYPE i,
        col_from TYPE i,
        col_to   TYPE i,
      END OF mty_merge .
    TYPES:
      mty_ts_merge TYPE SORTED TABLE OF mty_merge WITH UNIQUE KEY table_line .
    TYPES:
      ty_area TYPE c LENGTH 1 .

    CONSTANTS c_break_column TYPE zexcel_break VALUE 2 ##NO_TEXT.
    CONSTANTS c_break_none TYPE zexcel_break VALUE 0 ##NO_TEXT.
    CONSTANTS c_break_row TYPE zexcel_break VALUE 1 ##NO_TEXT.
    CONSTANTS:
      BEGIN OF c_area,
        whole   TYPE ty_area VALUE 'W',                     "#EC NOTEXT
        topleft TYPE ty_area VALUE 'T',                     "#EC NOTEXT
      END OF c_area .
    DATA excel TYPE REF TO zcl_excel READ-ONLY .
    DATA print_gridlines TYPE zexcel_print_gridlines READ-ONLY VALUE abap_false ##NO_TEXT.
    DATA sheet_content TYPE zexcel_t_cell_data .
    DATA sheet_setup TYPE REF TO zcl_excel_sheet_setup .
    DATA show_gridlines TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true ##NO_TEXT.
    DATA show_rowcolheaders TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true ##NO_TEXT.
    DATA tabcolor TYPE zexcel_s_tabcolor READ-ONLY .
    DATA column_formulas TYPE mty_th_column_formula READ-ONLY .
    CLASS-DATA:
      BEGIN OF c_messages READ-ONLY,
        formula_id_only_is_possible TYPE string,
        column_formula_id_not_found TYPE string,
        formula_not_in_this_table   TYPE string,
        formula_in_other_column     TYPE string,
      END OF c_messages .
    DATA mt_merged_cells TYPE mty_ts_merge READ-ONLY .
    DATA pane_top_left_cell TYPE string READ-ONLY.
    DATA sheetview_top_left_cell TYPE string READ-ONLY.

    METHODS add_comment
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS add_drawing
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS add_new_column
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS add_new_style_cond
      IMPORTING
        !ip_dimension_range  TYPE string DEFAULT 'A1'
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS add_new_data_validation
      RETURNING
        VALUE(eo_data_validation) TYPE REF TO zcl_excel_data_validation .
    METHODS add_new_range
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS add_new_row
      IMPORTING
        !ip_row       TYPE simple
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS bind_alv
      IMPORTING
        !io_alv      TYPE REF TO object
        !it_table    TYPE STANDARD TABLE
        !i_top       TYPE i DEFAULT 1
        !i_left      TYPE i DEFAULT 1
        !table_style TYPE zexcel_table_style OPTIONAL
        !i_table     TYPE abap_bool DEFAULT abap_true
      RAISING
        zcx_excel .
    METHODS bind_alv_ole2
      IMPORTING
        !i_document_url      TYPE ty_doc_url DEFAULT space
        !i_xls               TYPE c DEFAULT space
        !i_save_path         TYPE string
        !io_alv              TYPE REF TO cl_gui_alv_grid
        !it_listheader       TYPE slis_t_listheader OPTIONAL
        !i_top               TYPE i DEFAULT 1
        !i_left              TYPE i DEFAULT 1
        !i_columns_header    TYPE c DEFAULT 'X'
        !i_columns_autofit   TYPE c DEFAULT 'X'
        !i_format_col_header TYPE soi_format_item OPTIONAL
        !i_format_subtotal   TYPE soi_format_item OPTIONAL
        !i_format_total      TYPE soi_format_item OPTIONAL
      EXCEPTIONS
        miss_guide
        ex_transfer_kkblo_error
        fatal_error
        inv_data_range
        dim_mismatch_vkey
        dim_mismatch_sema
        error_in_sema .
    METHODS bind_table
      IMPORTING
        !ip_table               TYPE STANDARD TABLE
        !it_field_catalog       TYPE zexcel_t_fieldcatalog OPTIONAL
        !is_table_settings      TYPE zexcel_s_table_settings OPTIONAL
        VALUE(iv_default_descr) TYPE c OPTIONAL
        !iv_no_line_if_empty    TYPE abap_bool DEFAULT abap_false
        !ip_conv_exit_length    TYPE abap_bool DEFAULT abap_false
        !ip_conv_curr_amt_ext   TYPE abap_bool DEFAULT abap_false
      EXPORTING
        !es_table_settings      TYPE zexcel_s_table_settings
      RAISING
        zcx_excel .
    METHODS calculate_column_widths
      RAISING
        zcx_excel .
    METHODS change_area_style
      IMPORTING
        !ip_range         TYPE csequence OPTIONAL
        !ip_column_start  TYPE simple OPTIONAL
        !ip_column_end    TYPE simple OPTIONAL
        !ip_row           TYPE zexcel_cell_row OPTIONAL
        !ip_row_to        TYPE zexcel_cell_row OPTIONAL
        !ip_style_changer TYPE REF TO zif_excel_style_changer
      RAISING
        zcx_excel .
    METHODS change_cell_style
      IMPORTING
        !ip_columnrow                   TYPE csequence OPTIONAL
        !ip_column                      TYPE simple OPTIONAL
        !ip_row                         TYPE zexcel_cell_row OPTIONAL
        !ip_complete                    TYPE zexcel_s_cstyle_complete OPTIONAL
        !ip_xcomplete                   TYPE zexcel_s_cstylex_complete OPTIONAL
        !ip_font                        TYPE zexcel_s_cstyle_font OPTIONAL
        !ip_xfont                       TYPE zexcel_s_cstylex_font OPTIONAL
        !ip_fill                        TYPE zexcel_s_cstyle_fill OPTIONAL
        !ip_xfill                       TYPE zexcel_s_cstylex_fill OPTIONAL
        !ip_borders                     TYPE zexcel_s_cstyle_borders OPTIONAL
        !ip_xborders                    TYPE zexcel_s_cstylex_borders OPTIONAL
        !ip_alignment                   TYPE zexcel_s_cstyle_alignment OPTIONAL
        !ip_xalignment                  TYPE zexcel_s_cstylex_alignment OPTIONAL
        !ip_number_format_format_code   TYPE zexcel_number_format OPTIONAL
        !ip_protection                  TYPE zexcel_s_cstyle_protection OPTIONAL
        !ip_xprotection                 TYPE zexcel_s_cstylex_protection OPTIONAL
        !ip_font_bold                   TYPE flag OPTIONAL
        !ip_font_color                  TYPE zexcel_s_style_color OPTIONAL
        !ip_font_color_rgb              TYPE zexcel_style_color_argb OPTIONAL
        !ip_font_color_indexed          TYPE zexcel_style_color_indexed OPTIONAL
        !ip_font_color_theme            TYPE zexcel_style_color_theme OPTIONAL
        !ip_font_color_tint             TYPE zexcel_style_color_tint OPTIONAL
        !ip_font_family                 TYPE zexcel_style_font_family OPTIONAL
        !ip_font_italic                 TYPE flag OPTIONAL
        !ip_font_name                   TYPE zexcel_style_font_name OPTIONAL
        !ip_font_scheme                 TYPE zexcel_style_font_scheme OPTIONAL
        !ip_font_size                   TYPE zexcel_style_font_size OPTIONAL
        !ip_font_strikethrough          TYPE flag OPTIONAL
        !ip_font_underline              TYPE flag OPTIONAL
        !ip_font_underline_mode         TYPE zexcel_style_font_underline OPTIONAL
        !ip_fill_filltype               TYPE zexcel_fill_type OPTIONAL
        !ip_fill_rotation               TYPE zexcel_rotation OPTIONAL
        !ip_fill_fgcolor                TYPE zexcel_s_style_color OPTIONAL
        !ip_fill_fgcolor_rgb            TYPE zexcel_style_color_argb OPTIONAL
        !ip_fill_fgcolor_indexed        TYPE zexcel_style_color_indexed OPTIONAL
        !ip_fill_fgcolor_theme          TYPE zexcel_style_color_theme OPTIONAL
        !ip_fill_fgcolor_tint           TYPE zexcel_style_color_tint OPTIONAL
        !ip_fill_bgcolor                TYPE zexcel_s_style_color OPTIONAL
        !ip_fill_bgcolor_rgb            TYPE zexcel_style_color_argb OPTIONAL
        !ip_fill_bgcolor_indexed        TYPE zexcel_style_color_indexed OPTIONAL
        !ip_fill_bgcolor_theme          TYPE zexcel_style_color_theme OPTIONAL
        !ip_fill_bgcolor_tint           TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_allborders          TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_type          TYPE zexcel_s_gradient_type-type OPTIONAL
        !ip_fill_gradtype_degree        TYPE zexcel_s_gradient_type-degree OPTIONAL
        !ip_xborders_allborders         TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_diagonal            TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_bottom        TYPE zexcel_s_gradient_type-bottom OPTIONAL
        !ip_fill_gradtype_top           TYPE zexcel_s_gradient_type-top OPTIONAL
        !ip_xborders_diagonal           TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_diagonal_mode       TYPE zexcel_diagonal OPTIONAL
        !ip_fill_gradtype_right         TYPE zexcel_s_gradient_type-right OPTIONAL
        !ip_borders_down                TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_left          TYPE zexcel_s_gradient_type-left OPTIONAL
        !ip_fill_gradtype_position1     TYPE zexcel_s_gradient_type-position1 OPTIONAL
        !ip_xborders_down               TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_left                TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_position2     TYPE zexcel_s_gradient_type-position2 OPTIONAL
        !ip_fill_gradtype_position3     TYPE zexcel_s_gradient_type-position3 OPTIONAL
        !ip_xborders_left               TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_right               TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_xborders_right              TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_top                 TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_xborders_top                TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_alignment_horizontal        TYPE zexcel_alignment OPTIONAL
        !ip_alignment_vertical          TYPE zexcel_alignment OPTIONAL
        !ip_alignment_textrotation      TYPE zexcel_text_rotation OPTIONAL
        !ip_alignment_wraptext          TYPE flag OPTIONAL
        !ip_alignment_shrinktofit       TYPE flag OPTIONAL
        !ip_alignment_indent            TYPE zexcel_indent OPTIONAL
        !ip_protection_hidden           TYPE zexcel_cell_protection OPTIONAL
        !ip_protection_locked           TYPE zexcel_cell_protection OPTIONAL
        !ip_borders_allborders_style    TYPE zexcel_border OPTIONAL
        !ip_borders_allborders_color    TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_allbo_color_rgb     TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_allbo_color_indexed TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_allbo_color_theme   TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_allbo_color_tint    TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_diagonal_style      TYPE zexcel_border OPTIONAL
        !ip_borders_diagonal_color      TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_diagonal_color_rgb  TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_diagonal_color_inde TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_diagonal_color_them TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_diagonal_color_tint TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_down_style          TYPE zexcel_border OPTIONAL
        !ip_borders_down_color          TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_down_color_rgb      TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_down_color_indexed  TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_down_color_theme    TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_down_color_tint     TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_left_style          TYPE zexcel_border OPTIONAL
        !ip_borders_left_color          TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_left_color_rgb      TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_left_color_indexed  TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_left_color_theme    TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_left_color_tint     TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_right_style         TYPE zexcel_border OPTIONAL
        !ip_borders_right_color         TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_right_color_rgb     TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_right_color_indexed TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_right_color_theme   TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_right_color_tint    TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_top_style           TYPE zexcel_border OPTIONAL
        !ip_borders_top_color           TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_top_color_rgb       TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_top_color_indexed   TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_top_color_theme     TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_top_color_tint      TYPE zexcel_style_color_tint OPTIONAL
      RETURNING
        VALUE(ep_guid)                  TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    CLASS-METHODS class_constructor .
    METHODS constructor
      IMPORTING
        !ip_excel TYPE REF TO zcl_excel
        !ip_title TYPE zexcel_sheet_title OPTIONAL
      RAISING
        zcx_excel .
    METHODS delete_merge
      IMPORTING
        !ip_cell_column TYPE simple OPTIONAL
        !ip_cell_row    TYPE zexcel_cell_row OPTIONAL
      RAISING
        zcx_excel .
    METHODS delete_row_outline
      IMPORTING
        !iv_row_from TYPE i
        !iv_row_to   TYPE i
      RAISING
        zcx_excel .
    METHODS freeze_panes
      IMPORTING
        !ip_num_columns TYPE i OPTIONAL
        !ip_num_rows    TYPE i OPTIONAL
      RAISING
        zcx_excel .
    METHODS get_active_cell
      RETURNING
        VALUE(ep_active_cell) TYPE string
      RAISING
        zcx_excel .
    METHODS get_cell
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        !ep_value     TYPE zexcel_cell_value
        !ep_rc        TYPE sysubrc
        !ep_style     TYPE REF TO zcl_excel_style
        !ep_guid      TYPE zexcel_cell_style
        !ep_formula   TYPE zexcel_cell_formula
        !et_rtf       TYPE zexcel_t_rtf
      RAISING
        zcx_excel .
    METHODS get_column
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS get_columns
      RETURNING
        VALUE(eo_columns) TYPE REF TO zcl_excel_columns
      RAISING
        zcx_excel .
    METHODS get_columns_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator
      RAISING
        zcx_excel .
    METHODS get_style_cond_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_data_validations_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_data_validations_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_default_column
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS get_default_excel_date_format
      RETURNING
        VALUE(ep_default_excel_date_format) TYPE zexcel_number_format .
    METHODS get_default_excel_time_format
      RETURNING
        VALUE(ep_default_excel_time_format) TYPE zexcel_number_format .
    METHODS get_default_row
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_dimension_range
      RETURNING
        VALUE(ep_dimension_range) TYPE string
      RAISING
        zcx_excel .
    METHODS get_comments
      RETURNING
        VALUE(r_comments) TYPE REF TO zcl_excel_comments .
    METHODS get_drawings
      IMPORTING
        !ip_type          TYPE zexcel_drawing_type OPTIONAL
      RETURNING
        VALUE(r_drawings) TYPE REF TO zcl_excel_drawings .
    METHODS get_comments_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_drawings_iterator
      IMPORTING
        !ip_type           TYPE zexcel_drawing_type
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_freeze_cell
      EXPORTING
        !ep_row    TYPE zexcel_cell_row
        !ep_column TYPE zexcel_cell_column .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE sysuuid_x16 .
    METHODS get_highest_column
      RETURNING
        VALUE(r_highest_column) TYPE zexcel_cell_column
      RAISING
        zcx_excel .
    METHODS get_highest_row
      RETURNING
        VALUE(r_highest_row) TYPE int4
      RAISING
        zcx_excel .
    METHODS get_hyperlinks_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_hyperlinks_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_ignored_errors
      RETURNING
        VALUE(rt_ignored_errors) TYPE mty_th_ignored_errors .
    METHODS get_merge
      RETURNING
        VALUE(merge_range) TYPE string_table
      RAISING
        zcx_excel .
    METHODS get_pagebreaks
      RETURNING
        VALUE(ro_pagebreaks) TYPE REF TO zcl_excel_worksheet_pagebreaks
      RAISING
        zcx_excel .
    METHODS get_ranges_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_row
      IMPORTING
        !ip_row       TYPE int4
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_rows
      RETURNING
        VALUE(eo_rows) TYPE REF TO zcl_excel_rows .
    METHODS get_rows_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_row_outlines
      RETURNING
        VALUE(rt_row_outlines) TYPE mty_ts_outlines_row .
    METHODS get_style_cond
      IMPORTING
        !ip_guid             TYPE zexcel_cell_style
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS get_tabcolor
      RETURNING
        VALUE(ev_tabcolor) TYPE zexcel_s_tabcolor .
    METHODS get_tables_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_tables_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_title
      IMPORTING
        !ip_escaped     TYPE flag DEFAULT ''
      RETURNING
        VALUE(ep_title) TYPE zexcel_sheet_title .
    METHODS is_cell_merged
      IMPORTING
        !ip_column          TYPE simple
        !ip_row             TYPE zexcel_cell_row
      RETURNING
        VALUE(rp_is_merged) TYPE abap_bool
      RAISING
        zcx_excel .
    METHODS set_cell
      IMPORTING
        !ip_columnrow         TYPE csequence OPTIONAL
        !ip_column            TYPE simple OPTIONAL
        !ip_row               TYPE zexcel_cell_row OPTIONAL
        !ip_value             TYPE simple OPTIONAL
        !ip_formula           TYPE zexcel_cell_formula OPTIONAL
        !ip_style             TYPE any OPTIONAL
        !ip_hyperlink         TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type         TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type         TYPE abap_typekind OPTIONAL
        !ip_currency          TYPE waers_curc OPTIONAL
        !it_rtf               TYPE zexcel_t_rtf OPTIONAL
        !ip_column_formula_id TYPE mty_s_column_formula-id OPTIONAL
        !ip_conv_exit_length  TYPE abap_bool DEFAULT abap_false
      RAISING
        zcx_excel .
    METHODS set_cell_formula
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
        !ip_formula   TYPE zexcel_cell_formula
      RAISING
        zcx_excel .
    METHODS set_cell_style
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
        !ip_style     TYPE any
      RAISING
        zcx_excel .
    METHODS set_column_width
      IMPORTING
        !ip_column         TYPE simple
        !ip_width_fix      TYPE simple DEFAULT 0
        !ip_width_autosize TYPE flag DEFAULT 'X'
      RAISING
        zcx_excel .
    METHODS set_default_excel_date_format
      IMPORTING
        !ip_default_excel_date_format TYPE zexcel_number_format
      RAISING
        zcx_excel .
    METHODS set_ignored_errors
      IMPORTING
        !it_ignored_errors TYPE mty_th_ignored_errors .
    METHODS set_merge
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any OPTIONAL
        !ip_value        TYPE simple OPTIONAL          "added parameter
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL        "added parameter
      RAISING
        zcx_excel .
    METHODS set_pane_top_left_cell
      IMPORTING
        !iv_columnrow TYPE csequence
      RAISING
        zcx_excel.
    METHODS set_print_gridlines
      IMPORTING
        !i_print_gridlines TYPE zexcel_print_gridlines .
    METHODS set_row_height
      IMPORTING
        !ip_row        TYPE simple
        !ip_height_fix TYPE simple
      RAISING
        zcx_excel .
    METHODS set_row_outline
      IMPORTING
        !iv_row_from  TYPE i
        !iv_row_to    TYPE i
        !iv_collapsed TYPE abap_bool
      RAISING
        zcx_excel .
    METHODS set_sheetview_top_left_cell
      IMPORTING
        !iv_columnrow TYPE csequence
      RAISING
        zcx_excel.
    METHODS set_show_gridlines
      IMPORTING
        !i_show_gridlines TYPE zexcel_show_gridlines .
    METHODS set_show_rowcolheaders
      IMPORTING
        !i_show_rowcolheaders TYPE zexcel_show_rowcolheader .
    METHODS set_tabcolor
      IMPORTING
        !iv_tabcolor TYPE zexcel_s_tabcolor .
    METHODS set_table
      IMPORTING
        !ip_table           TYPE STANDARD TABLE
        !ip_hdr_style       TYPE any OPTIONAL
        !ip_body_style      TYPE any OPTIONAL
        !ip_table_title     TYPE string
        !ip_top_left_column TYPE zexcel_cell_column_alpha DEFAULT 'B'
        !ip_top_left_row    TYPE zexcel_cell_row DEFAULT 3
        !ip_transpose       TYPE abap_bool OPTIONAL
        !ip_no_header       TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_title
      IMPORTING
        !ip_title TYPE zexcel_sheet_title
      RAISING
        zcx_excel .
    METHODS get_table
      IMPORTING
        !iv_skipped_rows           TYPE int4 DEFAULT 0
        !iv_skipped_cols           TYPE int4 DEFAULT 0
        !iv_max_col                TYPE int4 OPTIONAL
        !iv_max_row                TYPE int4 OPTIONAL
        !iv_skip_bottom_empty_rows TYPE abap_bool DEFAULT abap_false
      EXPORTING
        !et_table                  TYPE STANDARD TABLE
      RAISING
        zcx_excel .
    METHODS set_merge_style
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_area_formula
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula
        !ip_merge        TYPE abap_bool OPTIONAL
        !ip_area         TYPE ty_area DEFAULT c_area-topleft
      RAISING
        zcx_excel .
    METHODS set_area_style
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any
        !ip_merge        TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_area
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_value        TYPE simple OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL
        !ip_style        TYPE any OPTIONAL
        !ip_hyperlink    TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type    TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type    TYPE abap_typekind OPTIONAL
        !ip_merge        TYPE abap_bool OPTIONAL
        !ip_area         TYPE ty_area DEFAULT c_area-topleft
      RAISING
        zcx_excel .
    METHODS get_header_footer_drawings
      RETURNING
        VALUE(rt_drawings) TYPE zexcel_t_drawings .
    METHODS set_area_hyperlink
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_url          TYPE string
        !ip_is_internal  TYPE abap_bool
      RAISING
        zcx_excel .
    "! excel upload, counterpart to BIND_TABLE
    "! @parameter it_field_catalog | field catalog, used to derive correct types
    "! @parameter iv_begin_row | starting row, by default 2 to skip header
    "! @parameter et_data | generic internal table, there may be conversion losses
    "! @parameter er_data | ref to internal table of string columns, to get raw data without conversion losses.
    METHODS convert_to_table
      IMPORTING
        !it_field_catalog TYPE zexcel_t_fieldcatalog OPTIONAL
        !iv_begin_row     TYPE int4 DEFAULT 2
        !iv_end_row       TYPE int4 DEFAULT 0
      EXPORTING
        !et_data          TYPE STANDARD TABLE
        !er_data          TYPE REF TO data
      RAISING
        zcx_excel .