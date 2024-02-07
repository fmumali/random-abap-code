  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
    TYPES ty_table_settings TYPE STANDARD TABLE OF zexcel_s_table_settings WITH DEFAULT KEY.
    DATA active_cell TYPE zexcel_s_cell_data .
    DATA charts TYPE REF TO zcl_excel_drawings .
    DATA columns TYPE REF TO zcl_excel_columns .
    DATA row_default TYPE REF TO zcl_excel_row .
    DATA column_default TYPE REF TO zcl_excel_column .
    DATA styles_cond TYPE REF TO zcl_excel_styles_cond .
    DATA data_validations TYPE REF TO zcl_excel_data_validations .
    DATA default_excel_date_format TYPE zexcel_number_format .
    DATA default_excel_time_format TYPE zexcel_number_format .
    DATA comments TYPE REF TO zcl_excel_comments .
    DATA drawings TYPE REF TO zcl_excel_drawings .
    DATA freeze_pane_cell_column TYPE zexcel_cell_column .
    DATA freeze_pane_cell_row TYPE zexcel_cell_row .
    DATA guid TYPE sysuuid_x16 .
    DATA hyperlinks TYPE REF TO zcl_excel_collection .
    DATA lower_cell TYPE zexcel_s_cell_data .
    DATA mo_pagebreaks TYPE REF TO zcl_excel_worksheet_pagebreaks .
    DATA mt_row_outlines TYPE mty_ts_outlines_row .
    DATA print_title_col_from TYPE zexcel_cell_column_alpha .
    DATA print_title_col_to TYPE zexcel_cell_column_alpha .
    DATA print_title_row_from TYPE zexcel_cell_row .
    DATA print_title_row_to TYPE zexcel_cell_row .
    DATA ranges TYPE REF TO zcl_excel_ranges .
    DATA rows TYPE REF TO zcl_excel_rows .
    DATA tables TYPE REF TO zcl_excel_collection .
    DATA title TYPE zexcel_sheet_title VALUE 'Worksheet'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA upper_cell TYPE zexcel_s_cell_data .
    DATA mt_ignored_errors TYPE mty_th_ignored_errors.
    DATA right_to_left TYPE abap_bool.

    METHODS calculate_cell_width
      IMPORTING
        !ip_column      TYPE simple
        !ip_row         TYPE zexcel_cell_row
      RETURNING
        VALUE(ep_width) TYPE f
      RAISING
        zcx_excel .
    CLASS-METHODS calculate_table_bottom_right
      IMPORTING
        ip_table         TYPE STANDARD TABLE
        it_field_catalog TYPE zexcel_t_fieldcatalog
      CHANGING
        cs_settings      TYPE zexcel_s_table_settings
      RAISING
        zcx_excel.
    CLASS-METHODS check_cell_column_formula
      IMPORTING
        it_column_formulas   TYPE mty_th_column_formula
        ip_column_formula_id TYPE mty_s_column_formula-id
        ip_formula           TYPE zexcel_cell_formula
        ip_value             TYPE simple
        ip_row               TYPE zexcel_cell_row
        ip_column            TYPE zexcel_cell_column
      RAISING
        zcx_excel.
    METHODS check_rtf
      IMPORTING
        !ip_value       TYPE simple
        VALUE(ip_style) TYPE zexcel_cell_style OPTIONAL
      CHANGING
        !ct_rtf         TYPE zexcel_t_rtf
      RAISING
        zcx_excel .
    CLASS-METHODS check_table_overlapping
      IMPORTING
        is_table_settings       TYPE zexcel_s_table_settings
        it_other_table_settings TYPE ty_table_settings
      RAISING
        zcx_excel.
    METHODS clear_initial_colorxfields
      IMPORTING
        is_color  TYPE zexcel_s_style_color
      CHANGING
        cs_xcolor TYPE zexcel_s_cstylex_color.
    METHODS create_data_conv_exit_length
      IMPORTING
        !ip_value       TYPE simple
      RETURNING
        VALUE(ep_value) TYPE REF TO data.
    METHODS generate_title
      RETURNING
        VALUE(ep_title) TYPE zexcel_sheet_title .
    METHODS get_value_type
      IMPORTING
        !ip_value      TYPE simple
      EXPORTING
        !ep_value      TYPE simple
        !ep_value_type TYPE abap_typekind .
    METHODS move_supplied_borders
      IMPORTING
        iv_border_supplied        TYPE abap_bool
        is_border                 TYPE zexcel_s_cstyle_border
        iv_xborder_supplied       TYPE abap_bool
        is_xborder                TYPE zexcel_s_cstylex_border
      CHANGING
        cs_complete_style_border  TYPE zexcel_s_cstyle_border
        cs_complete_stylex_border TYPE zexcel_s_cstylex_border.
    METHODS normalize_column_heading_texts
      IMPORTING
        iv_default_descr TYPE c
        it_field_catalog TYPE zexcel_t_fieldcatalog
      RETURNING
        VALUE(result)    TYPE zexcel_t_fieldcatalog.
    METHODS normalize_columnrow_parameter
      IMPORTING
        ip_columnrow TYPE csequence OPTIONAL
        ip_column    TYPE simple OPTIONAL
        ip_row       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        ep_column    TYPE zexcel_cell_column
        ep_row       TYPE zexcel_cell_row
      RAISING
        zcx_excel.
    METHODS normalize_range_parameter
      IMPORTING
        ip_range        TYPE csequence OPTIONAL
        ip_column_start TYPE simple OPTIONAL
        ip_column_end   TYPE simple OPTIONAL
        ip_row          TYPE zexcel_cell_row OPTIONAL
        ip_row_to       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        ep_column_start TYPE zexcel_cell_column
        ep_column_end   TYPE zexcel_cell_column
        ep_row          TYPE zexcel_cell_row
        ep_row_to       TYPE zexcel_cell_row
      RAISING
        zcx_excel.
    CLASS-METHODS normalize_style_parameter
      IMPORTING
        !ip_style_or_guid TYPE any
      RETURNING
        VALUE(rv_guid)    TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS print_title_set_range .
    METHODS update_dimension_range
      RAISING
        zcx_excel .