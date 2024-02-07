CLASS zcl_excel_theme DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    CONSTANTS c_theme_elements TYPE string VALUE 'themeElements'. "#EC NOTEXT
    CONSTANTS c_theme_object_def TYPE string VALUE 'objectDefaults'. "#EC NOTEXT
    CONSTANTS c_theme_extra_color TYPE string VALUE 'extraClrSchemeLst'. "#EC NOTEXT
    CONSTANTS c_theme_extlst TYPE string VALUE 'extLst'.    "#EC NOTEXT
    CONSTANTS c_theme TYPE string VALUE 'theme'.            "#EC NOTEXT
    CONSTANTS c_theme_name TYPE string VALUE 'name'.        "#EC NOTEXT
    CONSTANTS c_theme_xmlns TYPE string VALUE 'xmlns:a'.    "#EC NOTEXT
    CONSTANTS c_theme_prefix TYPE string VALUE 'a'.         "#EC NOTEXT
    CONSTANTS c_theme_prefix_write TYPE string VALUE 'a:'.  "#EC NOTEXT
    CONSTANTS c_theme_xmlns_val TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main'. "#EC NOTEXT

    METHODS constructor .
    METHODS read_theme
      IMPORTING
        VALUE(io_theme_xml) TYPE REF TO if_ixml_document .
    METHODS write_theme
      RETURNING
        VALUE(rv_xstring) TYPE xstring .
    METHODS set_color
      IMPORTING
        VALUE(iv_type)         TYPE string
        VALUE(iv_srgb)         TYPE zcl_excel_theme_color_scheme=>t_srgb OPTIONAL
        VALUE(iv_syscolorname) TYPE string OPTIONAL
        VALUE(iv_syscolorlast) TYPE zcl_excel_theme_color_scheme=>t_srgb OPTIONAL .
    METHODS set_color_scheme_name
      IMPORTING
        iv_name TYPE string .
    METHODS set_font
      IMPORTING
        iv_type     TYPE string
        iv_script   TYPE string
        iv_typeface TYPE string .
    METHODS set_latin_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_ea_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_cs_font
      IMPORTING
        VALUE(iv_type)        TYPE string
        VALUE(iv_typeface)    TYPE string
        VALUE(iv_panose)      TYPE string OPTIONAL
        VALUE(iv_pitchfamily) TYPE string OPTIONAL
        VALUE(iv_charset)     TYPE string OPTIONAL .
    METHODS set_font_scheme_name
      IMPORTING
        VALUE(iv_name) TYPE string .
    METHODS set_theme_name
      IMPORTING
        VALUE(iv_name) TYPE string .