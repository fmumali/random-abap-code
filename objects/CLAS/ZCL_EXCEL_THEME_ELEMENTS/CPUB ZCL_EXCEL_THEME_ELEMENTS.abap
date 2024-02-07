CLASS zcl_excel_theme_elements DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC

  GLOBAL FRIENDS zcl_excel_theme .

  PUBLIC SECTION.

    CONSTANTS c_color_scheme TYPE string VALUE 'clrScheme'. "#EC NOTEXT
    CONSTANTS c_font_scheme TYPE string VALUE 'fontScheme'. "#EC NOTEXT
    CONSTANTS c_fmt_scheme TYPE string VALUE 'fmtScheme'.   "#EC NOTEXT
    CONSTANTS c_theme_elements TYPE string VALUE 'themeElements'. "#EC NOTEXT

    METHODS constructor .
    METHODS load
      IMPORTING
        !io_elements TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .