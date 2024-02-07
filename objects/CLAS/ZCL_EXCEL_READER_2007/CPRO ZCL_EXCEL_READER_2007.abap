  PROTECTED SECTION.

    TYPES:
*"* protected components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!
      BEGIN OF t_relationship,
        id           TYPE string,
        type         TYPE string,
        target       TYPE string,
        targetmode   TYPE string,
        worksheet    TYPE REF TO zcl_excel_worksheet,
        sheetid      TYPE string,     "ins #235 - repeat rows/cols - needed to identify correct sheet
        localsheetid TYPE string,
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
        state   TYPE string,
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
    TYPES:
      BEGIN OF t_range,
        name         TYPE string,
        hidden       TYPE string,       "inserted with issue #235 because Autofilters didn't passthrough
        localsheetid TYPE string,       " issue #163
      END OF t_range .
    TYPES:
      t_fills   TYPE STANDARD TABLE OF REF TO zcl_excel_style_fill WITH NON-UNIQUE DEFAULT KEY .
    TYPES:
      t_borders TYPE STANDARD TABLE OF REF TO zcl_excel_style_borders WITH NON-UNIQUE DEFAULT KEY .
    TYPES:
      t_fonts   TYPE STANDARD TABLE OF REF TO zcl_excel_style_font WITH NON-UNIQUE DEFAULT KEY .
    TYPES:
      t_style_refs TYPE STANDARD TABLE OF REF TO zcl_excel_style WITH NON-UNIQUE DEFAULT KEY .
    TYPES:
      BEGIN OF t_color,
        indexed TYPE string,
        rgb     TYPE string,
        theme   TYPE string,
        tint    TYPE string,
      END OF t_color .
    TYPES:
      BEGIN OF t_rel_drawing,
        id          TYPE string,
        content     TYPE xstring,
        file_ext    TYPE string,
        content_xml TYPE REF TO if_ixml_document,
      END OF t_rel_drawing .
    TYPES:
      t_rel_drawings TYPE STANDARD TABLE OF t_rel_drawing WITH NON-UNIQUE DEFAULT KEY .
    TYPES:
      BEGIN OF gts_external_hyperlink,
        id     TYPE string,
        target TYPE string,
      END OF gts_external_hyperlink .
    TYPES:
      gtt_external_hyperlinks TYPE HASHED TABLE OF gts_external_hyperlink WITH UNIQUE KEY id .
    TYPES:
      BEGIN OF ty_ref_formulae,
        sheet   TYPE REF TO zcl_excel_worksheet,
        row     TYPE i,
        column  TYPE i,
        si      TYPE i,
        ref     TYPE string,
        formula TYPE string,
      END   OF ty_ref_formulae .
    TYPES:
      tyt_ref_formulae TYPE HASHED TABLE OF ty_ref_formulae WITH UNIQUE KEY sheet row column .
    TYPES:
      BEGIN OF t_shared_string,
        value TYPE string,
        rtf   TYPE zexcel_t_rtf,
      END OF t_shared_string .
    TYPES:
      t_shared_strings TYPE STANDARD TABLE OF t_shared_string WITH DEFAULT KEY .
    TYPES:
      BEGIN OF t_table,
        id     TYPE string,
        target TYPE string,
      END OF t_table .
    TYPES:
      t_tables TYPE HASHED TABLE OF t_table WITH UNIQUE KEY id .

    DATA shared_strings TYPE t_shared_strings .
    DATA styles TYPE t_style_refs .
    DATA mt_ref_formulae TYPE tyt_ref_formulae .
    DATA mt_dxf_styles TYPE zexcel_t_styles_cond_mapping .

    METHODS fill_row_outlines
      IMPORTING
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS get_from_zip_archive
      IMPORTING
        !i_filename      TYPE string
      RETURNING
        VALUE(r_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS get_ixml_from_zip_archive
      IMPORTING
        !i_filename     TYPE string
        !is_normalizing TYPE abap_bool DEFAULT 'X'
      RETURNING
        VALUE(r_ixml)   TYPE REF TO if_ixml_document
      RAISING
        zcx_excel .
    METHODS load_drawing_anchor
      IMPORTING
        !io_anchor_element   TYPE REF TO if_ixml_element
        !io_worksheet        TYPE REF TO zcl_excel_worksheet
        !it_related_drawings TYPE t_rel_drawings .
    METHODS load_shared_strings
      IMPORTING
        !ip_path TYPE string
      RAISING
        zcx_excel .
    METHODS load_styles
      IMPORTING
        !ip_path  TYPE string
        !ip_excel TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    METHODS load_dxf_styles
      IMPORTING
        !iv_path  TYPE string
        !io_excel TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    METHODS load_style_borders
      IMPORTING
        !ip_xml           TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_borders) TYPE t_borders .
    METHODS load_style_fills
      IMPORTING
        !ip_xml         TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_fills) TYPE t_fills .
    METHODS load_style_font
      IMPORTING
        !io_xml_element TYPE REF TO if_ixml_element
      RETURNING
        VALUE(ro_font)  TYPE REF TO zcl_excel_style_font .
    METHODS load_style_fonts
      IMPORTING
        !ip_xml         TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_fonts) TYPE t_fonts .
    METHODS load_style_num_formats
      IMPORTING
        !ip_xml               TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_num_formats) TYPE zcl_excel_style_number_format=>t_num_formats .
    METHODS load_workbook
      IMPORTING
        !iv_workbook_full_filename TYPE string
        !io_excel                  TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    METHODS load_worksheet
      IMPORTING
        !ip_path      TYPE string
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_worksheet_cond_format
      IMPORTING
        !io_ixml_worksheet TYPE REF TO if_ixml_document
        !io_worksheet      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_worksheet_cond_format_aa
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond.
    METHODS load_worksheet_cond_format_ci
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_cond_format_cs
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_cond_format_ex
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_cond_format_is
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_cond_format_db
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_cond_format_t10
      IMPORTING
        !io_ixml_rule  TYPE REF TO if_ixml_element
        !io_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS load_worksheet_drawing
      IMPORTING
        !ip_path      TYPE string
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_comments
      IMPORTING
        ip_path      TYPE string
        io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_worksheet_hyperlinks
      IMPORTING
        !io_ixml_worksheet      TYPE REF TO if_ixml_document
        !io_worksheet           TYPE REF TO zcl_excel_worksheet
        !it_external_hyperlinks TYPE gtt_external_hyperlinks
      RAISING
        zcx_excel .
    METHODS load_worksheet_ignored_errors
      IMPORTING
        !io_ixml_worksheet TYPE REF TO if_ixml_document
        !io_worksheet      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_worksheet_pagebreaks
      IMPORTING
        !io_ixml_worksheet TYPE REF TO if_ixml_document
        !io_worksheet      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS load_worksheet_autofilter
      IMPORTING
        io_ixml_worksheet TYPE REF TO if_ixml_document
        io_worksheet      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel.
    METHODS load_worksheet_pagemargins
      IMPORTING
        !io_ixml_worksheet TYPE REF TO if_ixml_document
        !io_worksheet      TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    "! <p class="shorttext synchronized" lang="en">Load worksheet tables</p>
    METHODS load_worksheet_tables
      IMPORTING
        io_ixml_worksheet TYPE REF TO if_ixml_document
        io_worksheet      TYPE REF TO zcl_excel_worksheet
        iv_dirname        TYPE string
        it_tables         TYPE t_tables
      RAISING
        zcx_excel .
    CLASS-METHODS resolve_path
      IMPORTING
        !ip_path         TYPE string
      RETURNING
        VALUE(rp_result) TYPE string .
    METHODS resolve_referenced_formulae .
    METHODS unescape_string_value
      IMPORTING
        i_value       TYPE string
      RETURNING
        VALUE(result) TYPE string.
    METHODS get_dxf_style_guid
      IMPORTING
        !io_ixml_dxf         TYPE REF TO if_ixml_element
        !io_excel            TYPE REF TO zcl_excel
      RETURNING
        VALUE(rv_style_guid) TYPE zexcel_cell_style .
    METHODS load_theme
      IMPORTING
        iv_path   TYPE string
        !ip_excel TYPE REF TO zcl_excel
      RAISING
        zcx_excel.
    METHODS provided_string_is_escaped
      IMPORTING
        !value            TYPE string
      RETURNING
        VALUE(is_escaped) TYPE abap_bool.

    CONSTANTS: BEGIN OF namespace,
                 x14ac            TYPE string VALUE 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
                 vba_project      TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', "#EC NEEDED     for future incorporation of XLSM-reader
                 c                TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                 a                TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main',
                 xdr              TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                 mc               TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                 r                TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                 chart            TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                 drawing          TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                 hyperlink        TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                 image            TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                 office_document  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                 printer_settings TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
                 shared_strings   TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                 styles           TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                 theme            TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                 worksheet        TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                 relationships    TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
                 core_properties  TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                 main             TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               END OF namespace.
