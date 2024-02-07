CLASS zcl_excel_ole DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    TYPES ty_doc_url TYPE c LENGTH 255.

    CLASS-METHODS bind_alv_ole2
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
