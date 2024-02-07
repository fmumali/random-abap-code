CLASS zcl_excel_graph_line DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_graph
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
*"* public components of class ZCL_EXCEL_GRAPH_LINE
*"* do not include other source files here!!!
      tv_type TYPE c LENGTH 5,
      BEGIN OF s_ax,
        axid          TYPE string,
        type          TYPE tv_type,
        orientation   TYPE string,
        delete        TYPE string,
        axpos         TYPE string,
        formatcode    TYPE string,
        sourcelinked  TYPE string,
        majortickmark TYPE string,
        minortickmark TYPE string,
        ticklblpos    TYPE string,
        crossax       TYPE string,
        crosses       TYPE string,
        auto          TYPE string,
        lblalgn       TYPE string,
        lbloffset     TYPE string,
        nomultilvllbl TYPE string,
        crossbetween  TYPE string,
      END OF s_ax .
    TYPES:
      t_ax TYPE STANDARD TABLE OF s_ax .

    DATA ns_groupingval TYPE string VALUE 'standard'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_varycolorsval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showlegendkeyval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showvalval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showcatnameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showsernameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showpercentval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showbubblesizeval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_markerval TYPE string VALUE '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_smoothval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA axes TYPE t_ax .
    CONSTANTS:
      c_valax TYPE c LENGTH 5 VALUE 'VALAX'.                "#EC NOTEXT
    CONSTANTS:
      c_catax TYPE c LENGTH 5 VALUE 'CATAX'.                "#EC NOTEXT
    DATA ns_legendposval TYPE string VALUE 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_overlayval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_symbol_auto TYPE string VALUE 'auto'.       "#EC NOTEXT
    CONSTANTS c_symbol_none TYPE string VALUE 'none'.       "#EC NOTEXT

    METHODS create_ax
      IMPORTING
        !ip_axid          TYPE string OPTIONAL
        !ip_type          TYPE tv_type
        !ip_orientation   TYPE string OPTIONAL
        !ip_delete        TYPE string OPTIONAL
        !ip_axpos         TYPE string OPTIONAL
        !ip_formatcode    TYPE string OPTIONAL
        !ip_sourcelinked  TYPE string OPTIONAL
        !ip_majortickmark TYPE string OPTIONAL
        !ip_minortickmark TYPE string OPTIONAL
        !ip_ticklblpos    TYPE string OPTIONAL
        !ip_crossax       TYPE string OPTIONAL
        !ip_crosses       TYPE string OPTIONAL
        !ip_auto          TYPE string OPTIONAL
        !ip_lblalgn       TYPE string OPTIONAL
        !ip_lbloffset     TYPE string OPTIONAL
        !ip_nomultilvllbl TYPE string OPTIONAL
        !ip_crossbetween  TYPE string OPTIONAL .
    METHODS set_show_legend_key
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_values
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_cat_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_ser_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_percent
      IMPORTING
        !ip_value TYPE c .
    METHODS set_varycolor
      IMPORTING
        !ip_value TYPE c .