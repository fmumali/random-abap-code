  PRIVATE SECTION.

    DATA zip TYPE REF TO lcl_zip_archive .
    DATA: gid TYPE i.

    METHODS create_zip_archive
      IMPORTING
        !i_xlsx_binary       TYPE xstring
        !i_use_alternate_zip TYPE seoclsname OPTIONAL
      RETURNING
        VALUE(e_zip)         TYPE REF TO lcl_zip_archive
      RAISING
        zcx_excel .
    METHODS read_from_applserver
      IMPORTING
        !i_filename         TYPE csequence
      RETURNING
        VALUE(r_excel_data) TYPE xstring
      RAISING
        zcx_excel.
    METHODS read_from_local_file
      IMPORTING
        !i_filename         TYPE csequence
      RETURNING
        VALUE(r_excel_data) TYPE xstring
      RAISING
        zcx_excel .