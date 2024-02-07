  PRIVATE SECTION.

    TYPES:
      BEGIN OF t_cell_content,
        datatype TYPE zexcel_cell_data_type,
        value    TYPE zexcel_cell_value,
        formula  TYPE zexcel_cell_formula,
        style    TYPE zexcel_cell_style,
      END OF t_cell_content .
    TYPES:
      BEGIN OF t_cell_coord,
        row    TYPE zexcel_cell_row,
        column TYPE zexcel_cell_column_alpha,
      END OF t_cell_coord .
    TYPES:
      BEGIN OF t_cell.
        INCLUDE TYPE t_cell_coord AS coord.
        INCLUDE TYPE t_cell_content AS content.
    TYPES: END OF t_cell .

    INTERFACE if_sxml_node LOAD .
    CONSTANTS c_end_of_stream TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_final. "#EC NOTEXT
    CONSTANTS c_element_open TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_element_open. "#EC NOTEXT
    CONSTANTS c_element_close TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_element_close. "#EC NOTEXT
    DATA:
      BEGIN OF gs_buffer_style,
        index TYPE i VALUE -1,
        guid  TYPE zexcel_cell_style,
      END OF gs_buffer_style .
    CONSTANTS c_attribute TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_attribute. "#EC NOTEXT
    CONSTANTS c_node_value TYPE if_sxml_node=>node_type VALUE if_sxml_node=>co_nt_value. "#EC NOTEXT

    METHODS skip_to
      IMPORTING
        !iv_element_name TYPE string
        !io_reader       TYPE REF TO if_sxml_reader
      RAISING
        lcx_not_found .
    METHODS fill_cell_from_attributes
      IMPORTING
        !io_reader     TYPE REF TO if_sxml_reader
      RETURNING
        VALUE(es_cell) TYPE t_cell
      RAISING
        lcx_not_found
        zcx_excel.
    METHODS read_shared_strings
      IMPORTING
        !io_reader               TYPE REF TO if_sxml_reader
      RETURNING
        VALUE(et_shared_strings) TYPE stringtab .
    METHODS get_cell_coord
      IMPORTING
        !iv_coord       TYPE string
      RETURNING
        VALUE(es_coord) TYPE t_cell_coord
      RAISING
        zcx_excel.
    METHODS put_cell_to_worksheet
      IMPORTING
        !io_worksheet TYPE REF TO zcl_excel_worksheet
        !is_cell      TYPE t_cell
      RAISING
        zcx_excel.
    METHODS get_shared_string
      IMPORTING
        !iv_index       TYPE any
      RETURNING
        VALUE(ev_value) TYPE string
      RAISING
        lcx_not_found .
    METHODS get_style
      IMPORTING
        !iv_index            TYPE any
      RETURNING
        VALUE(ev_style_guid) TYPE zexcel_cell_style
      RAISING
        lcx_not_found .
    METHODS read_worksheet_data
      IMPORTING
        !io_reader    TYPE REF TO if_sxml_reader
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        lcx_not_found
        zcx_excel .
    METHODS get_sxml_reader
      IMPORTING
        !iv_path         TYPE string
      RETURNING
        VALUE(eo_reader) TYPE REF TO if_sxml_reader
      RAISING
        zcx_excel .