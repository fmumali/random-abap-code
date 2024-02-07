  METHOD create_xl_sharedstrings.
*
* Redefinition using simple transformation instead of CL_IXML
*
** Constant node name

    TYPES:
      BEGIN OF ts_root,
        count        TYPE string,
        unique_count TYPE string,
      END OF ts_root.

    DATA:
      lo_iterator  TYPE REF TO zcl_excel_collection_iterator,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

    DATA:
      ls_root           TYPE ts_root,
      lt_cell_data      TYPE zexcel_t_cell_data_unsorted,
      ls_shared_string  TYPE zexcel_s_shared_string,
      lv_sytabix        TYPE i,
      lt_shared_strings TYPE TABLE OF zexcel_s_shared_string.

    FIELD-SYMBOLS:
      <sheet_content>     TYPE zexcel_s_cell_data.

**********************************************************************
* STEP 1: Collect strings from each worksheet

    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL " delete formula content
                           OR data_type    NE 's'.        " MvC: Only shared strings

    ls_root-count = lines( lt_cell_data ).
    CONDENSE ls_root-count.

    SORT lt_cell_data BY cell_value data_type.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data COMPARING cell_value data_type.

    ls_root-unique_count = lines( lt_cell_data ).
    CONDENSE ls_root-unique_count.

    LOOP AT lt_cell_data ASSIGNING <sheet_content>.

      lv_sytabix = sy-tabix - 1.
      ls_shared_string-string_no = lv_sytabix.
      ls_shared_string-string_value = <sheet_content>-cell_value.
      INSERT ls_shared_string INTO TABLE shared_strings.

    ENDLOOP.

**********************************************************************
* STEP 2: Create XML

    " Escape the string values, use of standard table in order to keep same sort order as in sorted table SHARED_STRINGS.
    CLEAR lt_shared_strings.
    LOOP AT shared_strings INTO ls_shared_string.
      ls_shared_string-string_value = escape_string_value( ls_shared_string-string_value ).
      APPEND ls_shared_string TO lt_shared_strings.
    ENDLOOP.

    CALL TRANSFORMATION zexcel_tr_shared_strings
      SOURCE root = ls_root
             shared_strings = lt_shared_strings
      OPTIONS xml_header = 'full'
      RESULT XML ep_content.


  ENDMETHOD.