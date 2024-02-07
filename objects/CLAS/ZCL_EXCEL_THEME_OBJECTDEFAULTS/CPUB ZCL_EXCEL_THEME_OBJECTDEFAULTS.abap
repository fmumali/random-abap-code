CLASS zcl_excel_theme_objectdefaults DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS load
      IMPORTING
        !io_object_def TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .