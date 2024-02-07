  PRIVATE SECTION.

    METHODS get_color
      IMPORTING
        !io_object      TYPE REF TO if_ixml_element
      RETURNING
        VALUE(rv_color) TYPE t_color .
    METHODS set_defaults .