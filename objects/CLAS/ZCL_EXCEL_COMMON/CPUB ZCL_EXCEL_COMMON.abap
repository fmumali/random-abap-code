CLASS zcl_excel_common DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS c_excel_baseline_date TYPE d VALUE '19000101'. "#EC NOTEXT
    CLASS-DATA c_excel_numfmt_offset TYPE int1 VALUE 164. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_excel_sheet_max_col TYPE int4 VALUE 16384.  "#EC NOTEXT
    CONSTANTS c_excel_sheet_min_col TYPE int4 VALUE 1.      "#EC NOTEXT
    CONSTANTS c_excel_sheet_max_row TYPE int4 VALUE 1048576. "#EC NOTEXT
    CONSTANTS c_excel_sheet_min_row TYPE int4 VALUE 1.      "#EC NOTEXT
    CLASS-DATA c_spras_en TYPE spras VALUE 'E'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    CLASS-DATA o_conv TYPE REF TO cl_abap_conv_out_ce .
    CONSTANTS c_excel_1900_leap_year TYPE d VALUE '19000228'. "#EC NOTEXT
    CLASS-DATA c_xlsx_file_filter TYPE string VALUE 'Excel Workbook (*.xlsx)|*.xlsx|'. "#EC NOTEXT .  .  .  .  .  .  . " .

    CLASS-METHODS class_constructor .
    CLASS-METHODS describe_structure
      IMPORTING
        !io_struct      TYPE REF TO cl_abap_structdescr
      RETURNING
        VALUE(rt_dfies) TYPE ddfields .
    CLASS-METHODS convert_column2alpha
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(ep_column) TYPE zexcel_cell_column_alpha
      RAISING
        zcx_excel .
    CLASS-METHODS convert_column2int
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(ep_column) TYPE zexcel_cell_column
      RAISING
        zcx_excel .
    CLASS-METHODS convert_column_a_row2columnrow
      IMPORTING
        !i_column          TYPE simple
        !i_row             TYPE zexcel_cell_row
      RETURNING
        VALUE(e_columnrow) TYPE string
      RAISING
        zcx_excel.
    CLASS-METHODS convert_columnrow2column_a_row
      IMPORTING
        !i_columnrow  TYPE clike
      EXPORTING
        !e_column     TYPE zexcel_cell_column_alpha
        !e_column_int TYPE zexcel_cell_column
        !e_row        TYPE zexcel_cell_row
      RAISING
        zcx_excel.
    CLASS-METHODS convert_range2column_a_row
      IMPORTING
        !i_range            TYPE clike
        !i_allow_1dim_range TYPE abap_bool DEFAULT abap_false
      EXPORTING
        !e_column_start     TYPE zexcel_cell_column_alpha
        !e_column_start_int TYPE zexcel_cell_column
        !e_column_end       TYPE zexcel_cell_column_alpha
        !e_column_end_int   TYPE zexcel_cell_column
        !e_row_start        TYPE zexcel_cell_row
        !e_row_end          TYPE zexcel_cell_row
        !e_sheet            TYPE clike
      RAISING
        zcx_excel .
    CLASS-METHODS convert_columnrow2column_o_row
      IMPORTING
        !i_columnrow TYPE clike
      EXPORTING
        !e_column    TYPE zexcel_cell_column_alpha
        !e_row       TYPE zexcel_cell_row .
    CLASS-METHODS clone_ixml_with_namespaces
      IMPORTING
        element       TYPE REF TO if_ixml_element
      RETURNING
        VALUE(result) TYPE REF TO if_ixml_element.
    CLASS-METHODS date_to_excel_string
      IMPORTING
        !ip_value       TYPE d
      RETURNING
        VALUE(ep_value) TYPE zexcel_cell_value .
    CLASS-METHODS encrypt_password
      IMPORTING
        !i_pwd                 TYPE zexcel_aes_password
      RETURNING
        VALUE(r_encrypted_pwd) TYPE zexcel_aes_password .
    CLASS-METHODS escape_string
      IMPORTING
        !ip_value               TYPE clike
      RETURNING
        VALUE(ep_escaped_value) TYPE string .
    CLASS-METHODS unescape_string
      IMPORTING
        !iv_escaped                TYPE clike
      RETURNING
        VALUE(ev_unescaped_string) TYPE string
      RAISING
        zcx_excel .
    CLASS-METHODS excel_string_to_date
      IMPORTING
        !ip_value       TYPE zexcel_cell_value
      RETURNING
        VALUE(ep_value) TYPE d
      RAISING
        zcx_excel .
    CLASS-METHODS excel_string_to_time
      IMPORTING
        !ip_value       TYPE zexcel_cell_value
      RETURNING
        VALUE(ep_value) TYPE t
      RAISING
        zcx_excel .
    CLASS-METHODS excel_string_to_number
      IMPORTING
        !ip_value       TYPE zexcel_cell_value
      RETURNING
        VALUE(ep_value) TYPE f
      RAISING
        zcx_excel .
    CLASS-METHODS get_fieldcatalog
      IMPORTING
        !ip_table              TYPE STANDARD TABLE
        !iv_hide_mandt         TYPE abap_bool DEFAULT abap_true
        !ip_conv_exit_length   TYPE abap_bool DEFAULT abap_false
      RETURNING
        VALUE(ep_fieldcatalog) TYPE zexcel_t_fieldcatalog .
    CLASS-METHODS number_to_excel_string
      IMPORTING
        VALUE(ip_value) TYPE numeric
        ip_currency     TYPE waers_curc OPTIONAL
      RETURNING
        VALUE(ep_value) TYPE zexcel_cell_value .
    CLASS-METHODS recursive_class_to_struct
      IMPORTING
        !i_source  TYPE any
      CHANGING
        !e_target  TYPE data
        !e_targetx TYPE data .
    CLASS-METHODS recursive_struct_to_class
      IMPORTING
        !i_source  TYPE data
        !i_sourcex TYPE data
      CHANGING
        !e_target  TYPE any .
    CLASS-METHODS time_to_excel_string
      IMPORTING
        !ip_value       TYPE t
      RETURNING
        VALUE(ep_value) TYPE zexcel_cell_value .
    TYPES: t_char10 TYPE c LENGTH 10.
    TYPES: t_char255 TYPE c LENGTH 255.
    CLASS-METHODS split_file
      IMPORTING
        !ip_file         TYPE t_char255
      EXPORTING
        !ep_file         TYPE t_char255
        !ep_extension    TYPE t_char10
        !ep_dotextension TYPE t_char10 .
    CLASS-METHODS calculate_cell_distance
      IMPORTING
        !iv_reference_cell TYPE clike
        !iv_current_cell   TYPE clike
      EXPORTING
        !ev_row_difference TYPE i
        !ev_col_difference TYPE i
      RAISING
        zcx_excel .
    CLASS-METHODS determine_resulting_formula
      IMPORTING
        !iv_reference_cell          TYPE clike
        !iv_reference_formula       TYPE clike
        !iv_current_cell            TYPE clike
      RETURNING
        VALUE(ev_resulting_formula) TYPE string
      RAISING
        zcx_excel .
    CLASS-METHODS shift_formula
      IMPORTING
        !iv_reference_formula       TYPE clike
        VALUE(iv_shift_cols)        TYPE i
        VALUE(iv_shift_rows)        TYPE i
      RETURNING
        VALUE(ev_resulting_formula) TYPE string
      RAISING
        zcx_excel .
    CLASS-METHODS is_cell_in_range
      IMPORTING
        !ip_column         TYPE simple
        !ip_row            TYPE zexcel_cell_row
        !ip_range          TYPE clike
      RETURNING
        VALUE(rp_in_range) TYPE abap_bool
      RAISING
        zcx_excel .
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_COMMON
*"* do not include other source files here!!!