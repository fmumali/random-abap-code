CLASS zcl_excel_style_cond DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    CLASS zcl_excel_style_conditional DEFINITION LOAD .

*"* public components of class ZCL_EXCEL_STYLE_COND
*"* do not include other source files here!!!
    TYPES tv_conditional_show_value TYPE c LENGTH 1.
    TYPES tv_textfunction TYPE string.
    TYPES: BEGIN OF ts_conditional_textfunction,
             text         TYPE string,
             textfunction TYPE tv_textfunction,
             cell_style   TYPE zexcel_cell_style,
           END OF ts_conditional_textfunction.
    CONSTANTS c_cfvo_type_formula TYPE zexcel_conditional_type VALUE 'formula'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_max TYPE zexcel_conditional_type VALUE 'max'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_min TYPE zexcel_conditional_type VALUE 'min'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_number TYPE zexcel_conditional_type VALUE 'num'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_percent TYPE zexcel_conditional_type VALUE 'percent'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_percentile TYPE zexcel_conditional_type VALUE 'percentile'. "#EC NOTEXT
    CONSTANTS c_iconset_3arrows TYPE zexcel_condition_rule_iconset VALUE '3Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_3arrowsgray TYPE zexcel_condition_rule_iconset VALUE '3ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_3flags TYPE zexcel_condition_rule_iconset VALUE '3Flags'. "#EC NOTEXT
    CONSTANTS c_iconset_3signs TYPE zexcel_condition_rule_iconset VALUE '3Signs'. "#EC NOTEXT
    CONSTANTS c_iconset_3symbols TYPE zexcel_condition_rule_iconset VALUE '3Symbols'. "#EC NOTEXT
    CONSTANTS c_iconset_3symbols2 TYPE zexcel_condition_rule_iconset VALUE '3Symbols2'. "#EC NOTEXT
    CONSTANTS c_iconset_3trafficlights TYPE zexcel_condition_rule_iconset VALUE ''. "#EC NOTEXT
    CONSTANTS c_iconset_3trafficlights2 TYPE zexcel_condition_rule_iconset VALUE '3TrafficLights2'. "#EC NOTEXT
    CONSTANTS c_iconset_4arrows TYPE zexcel_condition_rule_iconset VALUE '4Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_4arrowsgray TYPE zexcel_condition_rule_iconset VALUE '4ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_4rating TYPE zexcel_condition_rule_iconset VALUE '4Rating'. "#EC NOTEXT
    CONSTANTS c_iconset_4redtoblack TYPE zexcel_condition_rule_iconset VALUE '4RedToBlack'. "#EC NOTEXT
    CONSTANTS c_iconset_4trafficlights TYPE zexcel_condition_rule_iconset VALUE '4TrafficLights'. "#EC NOTEXT
    CONSTANTS c_iconset_5arrows TYPE zexcel_condition_rule_iconset VALUE '5Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_5arrowsgray TYPE zexcel_condition_rule_iconset VALUE '5ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_5quarters TYPE zexcel_condition_rule_iconset VALUE '5Quarters'. "#EC NOTEXT
    CONSTANTS c_iconset_5rating TYPE zexcel_condition_rule_iconset VALUE '5Rating'. "#EC NOTEXT
    CONSTANTS c_operator_beginswith TYPE zexcel_condition_operator VALUE 'beginsWith'. "#EC NOTEXT
    CONSTANTS c_operator_between TYPE zexcel_condition_operator VALUE 'between'. "#EC NOTEXT
    CONSTANTS c_operator_containstext TYPE zexcel_condition_operator VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_operator_endswith TYPE zexcel_condition_operator VALUE 'endsWith'. "#EC NOTEXT
    CONSTANTS c_operator_equal TYPE zexcel_condition_operator VALUE 'equal'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthan TYPE zexcel_condition_operator VALUE 'greaterThan'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthanorequal TYPE zexcel_condition_operator VALUE 'greaterThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_lessthan TYPE zexcel_condition_operator VALUE 'lessThan'. "#EC NOTEXT
    CONSTANTS c_operator_lessthanorequal TYPE zexcel_condition_operator VALUE 'lessThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_none TYPE zexcel_condition_operator VALUE ''. "#EC NOTEXT
    CONSTANTS c_operator_notcontains TYPE zexcel_condition_operator VALUE 'notContains'. "#EC NOTEXT
    CONSTANTS c_operator_notequal TYPE zexcel_condition_operator VALUE 'notEqual'. "#EC NOTEXT
    CONSTANTS c_textfunction_beginswith TYPE tv_textfunction VALUE 'beginsWith'. "#EC NOTEXT
    CONSTANTS c_textfunction_containstext TYPE tv_textfunction VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_textfunction_endswith TYPE tv_textfunction VALUE 'endsWith'. "#EC NOTEXT
    CONSTANTS c_textfunction_notcontains TYPE tv_textfunction VALUE 'notContains'. "#EC NOTEXT
    CONSTANTS c_rule_cellis TYPE zexcel_condition_rule VALUE 'cellIs'. "#EC NOTEXT
    CONSTANTS c_rule_containstext TYPE zexcel_condition_rule VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_rule_databar TYPE zexcel_condition_rule VALUE 'dataBar'. "#EC NOTEXT
    CONSTANTS c_rule_expression TYPE zexcel_condition_rule VALUE 'expression'. "#EC NOTEXT
    CONSTANTS c_rule_iconset TYPE zexcel_condition_rule VALUE 'iconSet'. "#EC NOTEXT
    CONSTANTS c_rule_colorscale TYPE zexcel_condition_rule VALUE 'colorScale'. "#EC NOTEXT
    CONSTANTS c_rule_none TYPE zexcel_condition_rule VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_rule_textfunction TYPE zexcel_condition_rule VALUE '<textfunction>'. "#EC NOTEXT
    CONSTANTS c_rule_top10 TYPE zexcel_condition_rule VALUE 'top10'. "#EC NOTEXT
    CONSTANTS c_rule_above_average TYPE zexcel_condition_rule VALUE 'aboveAverage'. "#EC NOTEXT
    CONSTANTS c_showvalue_false TYPE tv_conditional_show_value VALUE 0. "#EC NOTEXT
    CONSTANTS c_showvalue_true TYPE tv_conditional_show_value VALUE 1. "#EC NOTEXT
    DATA mode_cellis TYPE zexcel_conditional_cellis .
    DATA mode_textfunction TYPE ts_conditional_textfunction .
    DATA mode_colorscale TYPE zexcel_conditional_colorscale .
    DATA mode_databar TYPE zexcel_conditional_databar .
    DATA mode_expression TYPE zexcel_conditional_expression .
    DATA mode_iconset TYPE zexcel_conditional_iconset .
    DATA mode_top10 TYPE zexcel_conditional_top10 .
    DATA mode_above_average TYPE zexcel_conditional_above_avg .
    DATA priority TYPE zexcel_style_priority VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA rule TYPE zexcel_condition_rule .

    METHODS constructor
      IMPORTING
        !ip_guid            TYPE zexcel_cell_style OPTIONAL
        !ip_dimension_range TYPE string.
    METHODS get_dimension_range
      RETURNING
        VALUE(ep_dimension_range) TYPE string .
    METHODS set_range
      IMPORTING
        !ip_start_row    TYPE zexcel_cell_row
        !ip_start_column TYPE zexcel_cell_column_alpha
        !ip_stop_row     TYPE zexcel_cell_row
        !ip_stop_column  TYPE zexcel_cell_column_alpha
      RAISING
        zcx_excel .
    METHODS add_range
      IMPORTING
        !ip_start_row    TYPE zexcel_cell_row
        !ip_start_column TYPE zexcel_cell_column_alpha
        !ip_stop_row     TYPE zexcel_cell_row
        !ip_stop_column  TYPE zexcel_cell_column_alpha
      RAISING
        zcx_excel .
    CLASS-METHODS factory_cond_style_iconset
      IMPORTING
        !io_worksheet        TYPE REF TO zcl_excel_worksheet
        !iv_icon_type        TYPE zexcel_condition_rule_iconset DEFAULT c_iconset_3trafficlights2
        !iv_cfvo1_type       TYPE zexcel_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo1_value      TYPE zexcel_conditional_value OPTIONAL
        !iv_cfvo2_type       TYPE zexcel_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo2_value      TYPE zexcel_conditional_value OPTIONAL
        !iv_cfvo3_type       TYPE zexcel_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo3_value      TYPE zexcel_conditional_value OPTIONAL
        !iv_cfvo4_type       TYPE zexcel_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo4_value      TYPE zexcel_conditional_value OPTIONAL
        !iv_cfvo5_type       TYPE zexcel_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo5_value      TYPE zexcel_conditional_value OPTIONAL
        !iv_showvalue        TYPE tv_conditional_show_value DEFAULT zcl_excel_style_cond=>c_showvalue_true
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_cell_style .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!