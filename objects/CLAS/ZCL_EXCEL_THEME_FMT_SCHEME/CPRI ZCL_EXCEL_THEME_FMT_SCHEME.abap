  PRIVATE SECTION.

    DATA fmt_scheme TYPE REF TO if_ixml_element .

    METHODS get_default_fmt
      RETURNING
        VALUE(rv_string) TYPE string .

    METHODS parse_string
      IMPORTING iv_string      TYPE string
      RETURNING VALUE(ri_node) TYPE REF TO if_ixml_node.