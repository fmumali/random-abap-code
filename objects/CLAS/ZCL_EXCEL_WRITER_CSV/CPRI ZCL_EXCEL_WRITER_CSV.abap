  PRIVATE SECTION.

    DATA excel TYPE REF TO zcl_excel .
    CLASS-DATA delimiter TYPE c VALUE ';' ##NO_TEXT.
    CLASS-DATA enclosure TYPE c VALUE '"' ##NO_TEXT.
    CLASS-DATA:
      eol TYPE c LENGTH 2 VALUE cl_abap_char_utilities=>cr_lf ##NO_TEXT.
    CLASS-DATA worksheet_name TYPE zexcel_worksheets_name .
    CLASS-DATA worksheet_index TYPE zexcel_active_worksheet .

    METHODS create
      RETURNING
        VALUE(ep_excel) TYPE xstring
      RAISING
        zcx_excel.
    METHODS create_csv
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .