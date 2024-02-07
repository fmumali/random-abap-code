CLASS zcl_excel_style_fill DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_FILL
*"* do not include other source files here!!!
    CONSTANTS c_fill_none TYPE zexcel_fill_type VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_fill_solid TYPE zexcel_fill_type VALUE 'solid'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_linear TYPE zexcel_fill_type VALUE 'linear'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_path TYPE zexcel_fill_type VALUE 'path'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkdown TYPE zexcel_fill_type VALUE 'darkDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgray TYPE zexcel_fill_type VALUE 'darkGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkgrid TYPE zexcel_fill_type VALUE 'darkGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkhorizontal TYPE zexcel_fill_type VALUE 'darkHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darktrellis TYPE zexcel_fill_type VALUE 'darkTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkup TYPE zexcel_fill_type VALUE 'darkUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_darkvertical TYPE zexcel_fill_type VALUE 'darkVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray0625 TYPE zexcel_fill_type VALUE 'gray0625'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_gray125 TYPE zexcel_fill_type VALUE 'gray125'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightdown TYPE zexcel_fill_type VALUE 'lightDown'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgray TYPE zexcel_fill_type VALUE 'lightGray'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightgrid TYPE zexcel_fill_type VALUE 'lightGrid'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighthorizontal TYPE zexcel_fill_type VALUE 'lightHorizontal'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lighttrellis TYPE zexcel_fill_type VALUE 'lightTrellis'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightup TYPE zexcel_fill_type VALUE 'lightUp'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_lightvertical TYPE zexcel_fill_type VALUE 'lightVertical'. "#EC NOTEXT
    CONSTANTS c_fill_pattern_mediumgray TYPE zexcel_fill_type VALUE 'mediumGray'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal90 TYPE zexcel_fill_type VALUE 'horizontal90'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontal270 TYPE zexcel_fill_type VALUE 'horizontal270'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_horizontalb TYPE zexcel_fill_type VALUE 'horizontalb'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_vertical TYPE zexcel_fill_type VALUE 'vertical'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_fromcenter TYPE zexcel_fill_type VALUE 'fromCenter'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45 TYPE zexcel_fill_type VALUE 'diagonal45'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal45b TYPE zexcel_fill_type VALUE 'diagonal45b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135 TYPE zexcel_fill_type VALUE 'diagonal135'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_diagonal135b TYPE zexcel_fill_type VALUE 'diagonal135b'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlt TYPE zexcel_fill_type VALUE 'cornerLT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerlb TYPE zexcel_fill_type VALUE 'cornerLB'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrt TYPE zexcel_fill_type VALUE 'cornerRT'. "#EC NOTEXT
    CONSTANTS c_fill_gradient_cornerrb TYPE zexcel_fill_type VALUE 'cornerRB'. "#EC NOTEXT
    DATA gradtype TYPE zexcel_s_gradient_type .
    DATA filltype TYPE zexcel_fill_type .
    DATA rotation TYPE zexcel_rotation .
    DATA fgcolor TYPE zexcel_s_style_color .
    DATA bgcolor TYPE zexcel_s_style_color .

    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(es_fill) TYPE zexcel_s_style_fill .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!