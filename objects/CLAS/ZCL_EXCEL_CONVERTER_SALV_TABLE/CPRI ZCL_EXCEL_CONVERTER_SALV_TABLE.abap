  PRIVATE SECTION.

    METHODS load_data
      IMPORTING
        !io_salv  TYPE REF TO cl_salv_table
        !it_table TYPE STANDARD TABLE .
    METHODS is_intercept_data_active
      RETURNING
        VALUE(rv_result) TYPE abap_bool.