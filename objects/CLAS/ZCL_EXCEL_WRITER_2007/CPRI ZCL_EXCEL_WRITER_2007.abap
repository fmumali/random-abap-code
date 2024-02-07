  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    CONSTANTS c_off TYPE string VALUE '0'.                  "#EC NOTEXT
    CONSTANTS c_on TYPE string VALUE '1'.                   "#EC NOTEXT
    CONSTANTS c_xl_printersettings TYPE string VALUE 'xl/printerSettings/printerSettings#.bin'. "#EC NOTEXT
    TYPES: tv_charbool TYPE c LENGTH 5.

    METHODS add_1_val_child_node
      IMPORTING
        io_document   TYPE REF TO if_ixml_document
        io_parent     TYPE REF TO if_ixml_element
        iv_elem_name  TYPE string
        iv_attr_name  TYPE string
        iv_attr_value TYPE string.
    METHODS flag2bool
      IMPORTING
        !ip_flag          TYPE flag
      RETURNING
        VALUE(ep_boolean) TYPE tv_charbool  .