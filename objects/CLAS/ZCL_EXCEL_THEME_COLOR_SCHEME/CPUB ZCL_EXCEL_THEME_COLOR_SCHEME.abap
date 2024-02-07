CLASS zcl_excel_theme_color_scheme DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC

  GLOBAL FRIENDS zcl_excel_theme
                 zcl_excel_theme_elements .

  PUBLIC SECTION.

    TYPES t_srgb TYPE string .
    TYPES:
      BEGIN OF t_syscolor,
        val     TYPE string,
        lastclr TYPE t_srgb,
      END OF t_syscolor .
    TYPES:
      BEGIN OF t_color,
        srgb     TYPE t_srgb,
        syscolor TYPE t_syscolor,
      END OF t_color .

    CONSTANTS c_dark1 TYPE string VALUE 'dk1'.              "#EC NOTEXT
    CONSTANTS c_dark2 TYPE string VALUE 'dk2'.              "#EC NOTEXT
    CONSTANTS c_light1 TYPE string VALUE 'lt1'.             "#EC NOTEXT
    CONSTANTS c_light2 TYPE string VALUE 'lt2'.             "#EC NOTEXT
    CONSTANTS c_accent1 TYPE string VALUE 'accent1'.        "#EC NOTEXT
    CONSTANTS c_accent2 TYPE string VALUE 'accent2'.        "#EC NOTEXT
    CONSTANTS c_accent3 TYPE string VALUE 'accent3'.        "#EC NOTEXT
    CONSTANTS c_accent4 TYPE string VALUE 'accent4'.        "#EC NOTEXT
    CONSTANTS c_accent5 TYPE string VALUE 'accent5'.        "#EC NOTEXT
    CONSTANTS c_accent6 TYPE string VALUE 'accent6'.        "#EC NOTEXT
    CONSTANTS c_hlink TYPE string VALUE 'hlink'.            "#EC NOTEXT
    CONSTANTS c_folhlink TYPE string VALUE 'folHlink'.      "#EC NOTEXT
    CONSTANTS c_syscolor TYPE string VALUE 'sysClr'.        "#EC NOTEXT
    CONSTANTS c_srgbcolor TYPE string VALUE 'srgbClr'.      "#EC NOTEXT
    CONSTANTS c_val TYPE string VALUE 'val'.                "#EC NOTEXT
    CONSTANTS c_lastclr TYPE string VALUE 'lastClr'.        "#EC NOTEXT
    CONSTANTS c_name TYPE string VALUE 'name'.              "#EC NOTEXT
    CONSTANTS c_scheme TYPE string VALUE 'clrScheme'.       "#EC NOTEXT

    METHODS load
      IMPORTING
        !io_color_scheme TYPE REF TO if_ixml_element .
    METHODS set_color
      IMPORTING
        iv_type         TYPE string
        iv_srgb         TYPE t_srgb OPTIONAL
        iv_syscolorname TYPE string OPTIONAL
        iv_syscolorlast TYPE t_srgb .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
    METHODS constructor .
    METHODS set_name
      IMPORTING
        VALUE(iv_name) TYPE string .