  METHOD load_worksheet_ignored_errors.

    DATA: lo_ixml_ignored_errors TYPE REF TO if_ixml_node_collection,
          lo_ixml_ignored_error  TYPE REF TO if_ixml_element,
          lo_ixml_iterator       TYPE REF TO if_ixml_node_iterator,
          ls_ignored_error       TYPE zcl_excel_worksheet=>mty_s_ignored_errors,
          lt_ignored_errors      TYPE zcl_excel_worksheet=>mty_th_ignored_errors.

    DATA: BEGIN OF ls_raw_ignored_error,
            sqref              TYPE string,
            evalerror          TYPE string,
            twodigittextyear   TYPE string,
            numberstoredastext TYPE string,
            formula            TYPE string,
            formularange       TYPE string,
            unlockedformula    TYPE string,
            emptycellreference TYPE string,
            listdatavalidation TYPE string,
            calculatedcolumn   TYPE string,
          END OF ls_raw_ignored_error.

    CLEAR lt_ignored_errors.

    lo_ixml_ignored_errors =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'ignoredError' uri = namespace-main ).
    lo_ixml_iterator   =  lo_ixml_ignored_errors->create_iterator( ).
    lo_ixml_ignored_error  ?= lo_ixml_iterator->get_next( ).

    WHILE lo_ixml_ignored_error IS BOUND.

      fill_struct_from_attributes( EXPORTING
                                     ip_element   = lo_ixml_ignored_error
                                   CHANGING
                                     cp_structure = ls_raw_ignored_error ).

      CLEAR ls_ignored_error.
      ls_ignored_error-cell_coords = ls_raw_ignored_error-sqref.
      ls_ignored_error-eval_error = boolc( ls_raw_ignored_error-evalerror = '1' ).
      ls_ignored_error-two_digit_text_year = boolc( ls_raw_ignored_error-twodigittextyear = '1' ).
      ls_ignored_error-number_stored_as_text = boolc( ls_raw_ignored_error-numberstoredastext = '1' ).
      ls_ignored_error-formula = boolc( ls_raw_ignored_error-formula = '1' ).
      ls_ignored_error-formula_range = boolc( ls_raw_ignored_error-formularange = '1' ).
      ls_ignored_error-unlocked_formula = boolc( ls_raw_ignored_error-unlockedformula = '1' ).
      ls_ignored_error-empty_cell_reference = boolc( ls_raw_ignored_error-emptycellreference = '1' ).
      ls_ignored_error-list_data_validation = boolc( ls_raw_ignored_error-listdatavalidation = '1' ).
      ls_ignored_error-calculated_column  = boolc( ls_raw_ignored_error-calculatedcolumn = '1' ).

      INSERT ls_ignored_error INTO TABLE lt_ignored_errors.

      lo_ixml_ignored_error ?= lo_ixml_iterator->get_next( ).

    ENDWHILE.

    io_worksheet->set_ignored_errors( lt_ignored_errors ).

  ENDMETHOD.