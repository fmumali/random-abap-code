  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    TYPES: BEGIN OF mty_column_formula_used,
             id TYPE zexcel_s_cell_data-column_formula_id,
             si TYPE string,
             "! type: shared, etc.
             t  TYPE string,
           END OF mty_column_formula_used,
           mty_column_formulas_used TYPE HASHED TABLE OF mty_column_formula_used WITH UNIQUE KEY id.
    CONSTANTS c_content_types TYPE string VALUE '[Content_Types].xml'. "#EC NOTEXT
    CONSTANTS c_docprops_app TYPE string VALUE 'docProps/app.xml'. "#EC NOTEXT
    CONSTANTS c_docprops_core TYPE string VALUE 'docProps/core.xml'. "#EC NOTEXT
    CONSTANTS c_relationships TYPE string VALUE '_rels/.rels'. "#EC NOTEXT
    CONSTANTS c_xl_calcchain TYPE string VALUE 'xl/calcChain.xml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings TYPE string VALUE 'xl/drawings/drawing#.xml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings_rels TYPE string VALUE 'xl/drawings/_rels/drawing#.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_relationships TYPE string VALUE 'xl/_rels/workbook.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_sharedstrings TYPE string VALUE 'xl/sharedStrings.xml'. "#EC NOTEXT
    CONSTANTS c_xl_sheet TYPE string VALUE 'xl/worksheets/sheet#.xml'. "#EC NOTEXT
    CONSTANTS c_xl_sheet_rels TYPE string VALUE 'xl/worksheets/_rels/sheet#.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_styles TYPE string VALUE 'xl/styles.xml'. "#EC NOTEXT
    CONSTANTS c_xl_theme TYPE string VALUE 'xl/theme/theme1.xml'. "#EC NOTEXT
    CONSTANTS c_xl_workbook TYPE string VALUE 'xl/workbook.xml'. "#EC NOTEXT
    DATA excel TYPE REF TO zcl_excel .
    DATA shared_strings TYPE zexcel_t_shared_string .
    DATA styles_cond_mapping TYPE zexcel_t_styles_cond_mapping .
    DATA styles_mapping TYPE zexcel_t_styles_mapping .
    CONSTANTS c_xl_comments TYPE string VALUE 'xl/comments#.xml'. "#EC NOTEXT
    CONSTANTS cl_xl_drawing_for_comments TYPE string VALUE 'xl/drawings/vmlDrawing#.vml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings_vml_rels TYPE string VALUE 'xl/drawings/_rels/vmlDrawing#.vml.rels'. "#EC NOTEXT
    DATA ixml TYPE REF TO if_ixml.
    DATA control_characters TYPE string.

    METHODS create_xl_sheet_sheet_data
      IMPORTING
        !io_document                   TYPE REF TO if_ixml_document
        !io_worksheet                  TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(rv_ixml_sheet_data_root) TYPE REF TO if_ixml_element
      RAISING
        zcx_excel .
    METHODS add_further_data_to_zip
      IMPORTING
        !io_zip TYPE REF TO cl_abap_zip .
    METHODS create
      RETURNING
        VALUE(ep_excel) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_content_types
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_docprops_app
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_docprops_core
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_dxf_style
      IMPORTING
        !iv_cell_style    TYPE zexcel_cell_style
        !io_dxf_element   TYPE REF TO if_ixml_element
        !io_ixml_document TYPE REF TO if_ixml_document
        !it_cellxfs       TYPE zexcel_t_cellxfs
        !it_fonts         TYPE zexcel_t_style_font
        !it_fills         TYPE zexcel_t_style_fill
      CHANGING
        !cv_dfx_count     TYPE i .
    METHODS create_relationships
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_charts
      IMPORTING
        !io_drawing       TYPE REF TO zcl_excel_drawing
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_comments
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawing_anchor
      IMPORTING
        !io_drawing      TYPE REF TO zcl_excel_drawing
        !io_document     TYPE REF TO if_ixml_document
        !ip_index        TYPE i
      RETURNING
        VALUE(ep_anchor) TYPE REF TO if_ixml_element .
    METHODS create_xl_drawing_for_comments
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_relationships
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_sharedstrings
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_sheet
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
        !iv_active        TYPE flag DEFAULT ''
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_sheet_ignored_errors
      IMPORTING
        io_worksheet    TYPE REF TO zcl_excel_worksheet
        io_document     TYPE REF TO if_ixml_document
        io_element_root TYPE REF TO if_ixml_element.
    METHODS create_xl_sheet_pagebreaks
      IMPORTING
        !io_document  TYPE REF TO if_ixml_document
        !io_parent    TYPE REF TO if_ixml_element
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS create_xl_sheet_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
        !iv_drawing_index TYPE i
        !iv_comment_index TYPE i
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_styles
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_styles_color_node
      IMPORTING
        !io_document        TYPE REF TO if_ixml_document
        !io_parent          TYPE REF TO if_ixml_element
        !iv_color_elem_name TYPE string DEFAULT 'color'
        !is_color           TYPE zexcel_s_style_color .
    METHODS create_xl_styles_font_node
      IMPORTING
        !io_document TYPE REF TO if_ixml_document
        !io_parent   TYPE REF TO if_ixml_element
        !is_font     TYPE zexcel_s_style_font
        !iv_use_rtf  TYPE abap_bool DEFAULT abap_false .
    METHODS create_xl_table
      IMPORTING
        !io_table         TYPE REF TO zcl_excel_table
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_theme
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_workbook
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS get_shared_string_index
      IMPORTING
        !ip_cell_value  TYPE zexcel_cell_value
        !it_rtf         TYPE zexcel_t_rtf OPTIONAL
      RETURNING
        VALUE(ep_index) TYPE int4 .
    METHODS create_xl_drawings_vml
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS set_vml_string
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS create_xl_drawings_vml_rels
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS escape_string_value
      IMPORTING
        !iv_value     TYPE zexcel_cell_value
      RETURNING
        VALUE(result) TYPE zexcel_cell_value.
    METHODS set_vml_shape_footer
      IMPORTING
        !is_footer        TYPE zexcel_s_worksheet_head_foot
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS set_vml_shape_header
      IMPORTING
        !is_header        TYPE zexcel_s_worksheet_head_foot
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS create_xl_drawing_for_hdft_im
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings_hdft_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xml_document
      RETURNING
        VALUE(ro_document) TYPE REF TO if_ixml_document.
    METHODS render_xml_document
      IMPORTING
        io_document       TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_content) TYPE xstring.
    METHODS create_xl_sheet_column_formula
      IMPORTING
        io_document             TYPE REF TO if_ixml_document
        it_column_formulas      TYPE zcl_excel_worksheet=>mty_th_column_formula
        is_sheet_content        TYPE zexcel_s_cell_data
      EXPORTING
        eo_element              TYPE REF TO if_ixml_element
      CHANGING
        ct_column_formulas_used TYPE mty_column_formulas_used
        cv_si                   TYPE i
      RAISING
        zcx_excel.
    METHODS is_formula_shareable
      IMPORTING
        ip_formula          TYPE string
      RETURNING
        VALUE(ep_shareable) TYPE abap_bool
      RAISING
        zcx_excel.