  PRIVATE SECTION.
    CLASS-METHODS get_subclasses
      IMPORTING
        is_clskey  TYPE seoclskey
      CHANGING
        ct_classes TYPE seor_implementing_keys.

    DATA wo_excel TYPE REF TO zcl_excel .
    DATA wo_worksheet TYPE REF TO zcl_excel_worksheet .
    DATA wo_autofilter TYPE REF TO zcl_excel_autofilter .
    DATA wo_table TYPE REF TO data .
    DATA wo_data TYPE REF TO data .
    DATA wt_fieldcatalog TYPE zexcel_t_converter_fcat .
    DATA ws_layout TYPE zexcel_s_converter_layo .
    DATA wt_colors TYPE zexcel_t_converter_col .
    DATA wt_filter TYPE zexcel_t_converter_fil .
    CLASS-DATA wt_objects TYPE tt_alv_types .
    CLASS-DATA w_fcount TYPE n LENGTH 3 .
    DATA wt_sort_values TYPE tt_sort_values .
    DATA wt_subtotal_rows TYPE tt_subtotal_rows .
    DATA wt_styles TYPE tt_styles .
    CONSTANTS c_type_hdr TYPE c VALUE 'H'.                  "#EC NOTEXT
    CONSTANTS c_type_str TYPE c VALUE 'P'.                  "#EC NOTEXT
    CONSTANTS c_type_nor TYPE c VALUE 'N'.                  "#EC NOTEXT
    CONSTANTS c_type_sub TYPE c VALUE 'S'.                  "#EC NOTEXT
    CONSTANTS c_type_tot TYPE c VALUE 'T'.                  "#EC NOTEXT
    DATA wt_color_styles TYPE tt_color_styles .
    CLASS-DATA ws_option TYPE zexcel_s_converter_option .
    CLASS-DATA ws_indx TYPE indx .

    CLASS-METHODS init_option .
    CLASS-METHODS get_alv_converters.
    METHODS bind_table
      IMPORTING
        !i_style_table      TYPE zexcel_table_style
      RETURNING
        VALUE(r_freeze_col) TYPE int1
      RAISING
        zcx_excel .
    METHODS bind_cells
      RETURNING
        VALUE(r_freeze_col) TYPE int1
      RAISING
        zcx_excel .
    METHODS clean_fieldcatalog .
    METHODS create_color_style
      IMPORTING
        !i_style        TYPE zexcel_cell_style
        !is_colors      TYPE zexcel_s_converter_col
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_formular_subtotal
      IMPORTING
        !i_row_int_start   TYPE zexcel_cell_row
        !i_row_int_end     TYPE zexcel_cell_row
        !i_column          TYPE zexcel_cell_column_alpha
        !i_totals_function TYPE zexcel_table_totals_function
      RETURNING
        VALUE(r_formula)   TYPE string .
    METHODS create_formular_total
      IMPORTING
        !i_row_int         TYPE zexcel_cell_row
        !i_column          TYPE zexcel_cell_column_alpha
        !i_totals_function TYPE zexcel_table_totals_function
      RETURNING
        VALUE(r_formula)   TYPE string .
    METHODS create_style_hdr
      IMPORTING
        !i_alignment    TYPE zexcel_alignment OPTIONAL
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_style_normal
      IMPORTING
        !i_alignment    TYPE zexcel_alignment OPTIONAL
        !i_inttype      TYPE abap_typekind OPTIONAL
        !i_decimals     TYPE int1 OPTIONAL
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_style_stripped
      IMPORTING
        !i_alignment    TYPE zexcel_alignment OPTIONAL
        !i_inttype      TYPE abap_typekind OPTIONAL
        !i_decimals     TYPE int1 OPTIONAL
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_style_subtotal
      IMPORTING
        !i_alignment    TYPE zexcel_alignment OPTIONAL
        !i_inttype      TYPE abap_typekind OPTIONAL
        !i_decimals     TYPE int1 OPTIONAL
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_style_total
      IMPORTING
        !i_alignment    TYPE zexcel_alignment OPTIONAL
        !i_inttype      TYPE abap_typekind OPTIONAL
        !i_decimals     TYPE int1 OPTIONAL
      RETURNING
        VALUE(ro_style) TYPE REF TO zcl_excel_style .
    METHODS create_table .
    METHODS create_text_subtotal
      IMPORTING
        !i_value           TYPE any
        !i_totals_function TYPE zexcel_table_totals_function
      RETURNING
        VALUE(r_text)      TYPE string .
    METHODS create_worksheet
      IMPORTING
        !i_table       TYPE flag DEFAULT 'X'
        !i_style_table TYPE zexcel_table_style
      RAISING
        zcx_excel .
    METHODS execute_converter
      IMPORTING
        !io_object TYPE REF TO object
        !it_table  TYPE STANDARD TABLE
      RAISING
        zcx_excel .
    METHODS get_color_style
      IMPORTING
        !i_row         TYPE zexcel_cell_row
        !i_fieldname   TYPE fieldname
        !i_style       TYPE zexcel_cell_style
      RETURNING
        VALUE(r_style) TYPE zexcel_cell_style .
    METHODS get_function_number
      IMPORTING
        !i_totals_function       TYPE zexcel_table_totals_function
      RETURNING
        VALUE(r_function_number) TYPE int1 .
    METHODS get_style
      IMPORTING
        !i_type        TYPE ty_style_type
        !i_alignment   TYPE zexcel_alignment DEFAULT space
        !i_inttype     TYPE abap_typekind DEFAULT space
        !i_decimals    TYPE int1 DEFAULT 0
      RETURNING
        VALUE(r_style) TYPE zexcel_cell_style .
    METHODS loop_normal
      IMPORTING
        !i_row_int          TYPE zexcel_cell_row
        !i_col_int          TYPE zexcel_cell_column
      RETURNING
        VALUE(r_freeze_col) TYPE int1
      RAISING
        zcx_excel .
    METHODS loop_subtotal
      IMPORTING
        !i_row_int          TYPE zexcel_cell_row
        !i_col_int          TYPE zexcel_cell_column
      RETURNING
        VALUE(r_freeze_col) TYPE int1
      RAISING
        zcx_excel .
    METHODS set_autofilter_area .
    METHODS set_cell_format
      IMPORTING
        !i_inttype      TYPE abap_typekind
        !i_decimals     TYPE int1
      RETURNING
        VALUE(r_format) TYPE zexcel_number_format .
    METHODS set_fieldcatalog .