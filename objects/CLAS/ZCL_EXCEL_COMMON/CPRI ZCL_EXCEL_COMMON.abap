  PRIVATE SECTION.

    CLASS-DATA c_excel_col_module TYPE int2 VALUE 64. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    CLASS-DATA sv_prev_in1  TYPE zexcel_cell_column.
    CLASS-DATA sv_prev_out1 TYPE zexcel_cell_column_alpha.
    CLASS-DATA sv_prev_in2  TYPE c LENGTH 10.
    CLASS-DATA sv_prev_out2 TYPE zexcel_cell_column.
    CLASS-METHODS structure_case
      IMPORTING
        !is_component  TYPE abap_componentdescr
      CHANGING
        !xt_components TYPE abap_component_tab .
    CLASS-METHODS structure_recursive
      IMPORTING
        !is_component        TYPE abap_componentdescr
      RETURNING
        VALUE(rt_components) TYPE abap_component_tab .
    TYPES ty_char1 TYPE c LENGTH 1.
    CLASS-METHODS char2hex
      IMPORTING
        !i_char      TYPE ty_char1
      RETURNING
        VALUE(r_hex) TYPE zexcel_pwd_hash .
    CLASS-METHODS shl01
      IMPORTING
        !i_pwd_hash       TYPE zexcel_pwd_hash
      RETURNING
        VALUE(r_pwd_hash) TYPE zexcel_pwd_hash .
    CLASS-METHODS shr14
      IMPORTING
        !i_pwd_hash       TYPE zexcel_pwd_hash
      RETURNING
        VALUE(r_pwd_hash) TYPE zexcel_pwd_hash .