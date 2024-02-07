CLASS zcl_excel_theme_font_scheme DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF t_font,
        script   TYPE string,
        typeface TYPE string,
      END OF t_font .
    TYPES:
      tt_font TYPE SORTED TABLE OF t_font WITH UNIQUE KEY script .
    TYPES:
      BEGIN OF t_fonttype,
        typeface    TYPE string,
        panose      TYPE string,
        pitchfamily TYPE string,
        charset     TYPE string,
      END OF t_fonttype .
    TYPES:
      BEGIN OF t_fonts,
        latin TYPE t_fonttype,
        ea    TYPE t_fonttype,
        cs    TYPE t_fonttype,
        fonts TYPE tt_font,
      END OF t_fonts .
    TYPES:
      BEGIN OF t_scheme,
        name  TYPE string,
        major TYPE t_fonts,
        minor TYPE t_fonts,
      END OF t_scheme .

    CONSTANTS c_name TYPE string VALUE 'name'.              "#EC NOTEXT
    CONSTANTS c_scheme TYPE string VALUE 'fontScheme'.      "#EC NOTEXT
    CONSTANTS c_major TYPE string VALUE 'majorFont'.        "#EC NOTEXT
    CONSTANTS c_minor TYPE string VALUE 'minorFont'.        "#EC NOTEXT
    CONSTANTS c_font TYPE string VALUE 'font'.              "#EC NOTEXT
    CONSTANTS c_latin TYPE string VALUE 'latin'.            "#EC NOTEXT
    CONSTANTS c_ea TYPE string VALUE 'ea'.                  "#EC NOTEXT
    CONSTANTS c_cs TYPE string VALUE 'cs'.                  "#EC NOTEXT
    CONSTANTS c_typeface TYPE string VALUE 'typeface'.      "#EC NOTEXT
    CONSTANTS c_panose TYPE string VALUE 'panose'.          "#EC NOTEXT
    CONSTANTS c_pitchfamily TYPE string VALUE 'pitchFamily'. "#EC NOTEXT
    CONSTANTS c_charset TYPE string VALUE 'charset'.        "#EC NOTEXT
    CONSTANTS c_script TYPE string VALUE 'script'.          "#EC NOTEXT

    METHODS load
      IMPORTING
        !io_font_scheme TYPE REF TO if_ixml_element .
    METHODS set_name
      IMPORTING
        iv_name TYPE string .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
    METHODS modify_font
      IMPORTING
        VALUE(iv_type)     TYPE string
        VALUE(iv_script)   TYPE string
        VALUE(iv_typeface) TYPE string .
    METHODS modify_latin_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS modify_ea_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS modify_cs_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS constructor .