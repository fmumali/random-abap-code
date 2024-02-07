CLASS zcl_excel_drawing DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    CONSTANTS c_graph_pie TYPE zexcel_graph_type VALUE 1.   "#EC NOTEXT
    CONSTANTS c_graph_line TYPE zexcel_graph_type VALUE 2.  "#EC NOTEXT
    CONSTANTS c_graph_bars TYPE zexcel_graph_type VALUE 0.  "#EC NOTEXT
    DATA graph_type TYPE zexcel_graph_type .
    DATA title TYPE string VALUE 'image1.jpg'.              "#EC NOTEXT
    CONSTANTS type_image TYPE zexcel_drawing_type VALUE 'image'. "#EC NOTEXT
    CONSTANTS type_chart TYPE zexcel_drawing_type VALUE 'chart'. "#EC NOTEXT
    CONSTANTS anchor_absolute TYPE zexcel_drawing_anchor VALUE 'ABS'. "#EC NOTEXT
    CONSTANTS anchor_one_cell TYPE zexcel_drawing_anchor VALUE 'ONE'. "#EC NOTEXT
    CONSTANTS anchor_two_cell TYPE zexcel_drawing_anchor VALUE 'TWO'. "#EC NOTEXT
    DATA graph TYPE REF TO zcl_excel_graph .
    CONSTANTS c_media_type_bmp TYPE string VALUE 'bmp'.     "#EC NOTEXT
    CONSTANTS c_media_type_xml TYPE string VALUE 'xml'.     "#EC NOTEXT
    CONSTANTS c_media_type_jpg TYPE string VALUE 'jpg'.     "#EC NOTEXT
    CONSTANTS type_image_header_footer TYPE zexcel_drawing_type VALUE 'hd_ft'. "#EC NOTEXT
    CONSTANTS: BEGIN OF namespace,
                 c   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                 a   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main',
                 r   TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                 mc  TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                 c14 TYPE string VALUE 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
               END OF namespace.

    METHODS constructor
      IMPORTING
        !ip_type  TYPE zexcel_drawing_type DEFAULT zcl_excel_drawing=>type_image
        !ip_title TYPE clike OPTIONAL .
    METHODS create_media_name
      IMPORTING
        !ip_index TYPE i .
    METHODS get_from_col
      RETURNING
        VALUE(r_from_col) TYPE zexcel_cell_column .
    METHODS get_from_row
      RETURNING
        VALUE(r_from_row) TYPE zexcel_cell_row .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_guid .
    METHODS get_height_emu_str
      RETURNING
        VALUE(r_height) TYPE string .
    METHODS get_media
      RETURNING
        VALUE(r_media) TYPE xstring .
    METHODS get_media_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_media_type
      RETURNING
        VALUE(r_type) TYPE string .
    METHODS get_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_to_col
      RETURNING
        VALUE(r_to_col) TYPE zexcel_cell_column .
    METHODS get_to_row
      RETURNING
        VALUE(r_to_row) TYPE zexcel_cell_row .
    METHODS get_width_emu_str
      RETURNING
        VALUE(r_width) TYPE string .
    CLASS-METHODS pixel2emu
      IMPORTING
        !ip_pixel    TYPE int4
        !ip_dpi      TYPE int2 OPTIONAL
      RETURNING
        VALUE(r_emu) TYPE string .
    CLASS-METHODS emu2pixel
      IMPORTING
        !ip_emu        TYPE int4
        !ip_dpi        TYPE int2 OPTIONAL
      RETURNING
        VALUE(r_pixel) TYPE int4 .
    METHODS set_media
      IMPORTING
        !ip_media      TYPE xstring OPTIONAL
        !ip_media_type TYPE string
        !ip_width      TYPE int4 DEFAULT 0
        !ip_height     TYPE int4 DEFAULT 0 .
    METHODS set_media_mime
      IMPORTING
        !ip_io     TYPE skwf_io
        !ip_width  TYPE int4
        !ip_height TYPE int4 .
    METHODS set_media_www
      IMPORTING
        !ip_key    TYPE wwwdatatab
        !ip_width  TYPE int4
        !ip_height TYPE int4 .
    METHODS set_position
      IMPORTING
        !ip_from_row TYPE zexcel_cell_row
        !ip_from_col TYPE zexcel_cell_column_alpha
        !ip_rowoff   TYPE int4 OPTIONAL
        !ip_coloff   TYPE int4 OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_position2
      IMPORTING
        !ip_from   TYPE zexcel_drawing_location
        !ip_to     TYPE zexcel_drawing_location
        !ip_anchor TYPE zexcel_drawing_anchor OPTIONAL .
    METHODS get_position
      RETURNING
        VALUE(rp_position) TYPE zexcel_drawing_position .
    METHODS get_type
      RETURNING
        VALUE(rp_type) TYPE zexcel_drawing_type .
    METHODS get_index
      RETURNING
        VALUE(rp_index) TYPE string .
    METHODS load_chart_attributes
      IMPORTING
        VALUE(ip_chart) TYPE REF TO if_ixml_document .