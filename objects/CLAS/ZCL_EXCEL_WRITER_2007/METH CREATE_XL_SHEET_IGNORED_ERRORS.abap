  METHOD create_xl_sheet_ignored_errors.
    DATA: lo_element        TYPE REF TO if_ixml_element,
          lo_element2       TYPE REF TO if_ixml_element,
          lt_ignored_errors TYPE zcl_excel_worksheet=>mty_th_ignored_errors.
    FIELD-SYMBOLS: <ls_ignored_errors> TYPE zcl_excel_worksheet=>mty_s_ignored_errors.

    lt_ignored_errors = io_worksheet->get_ignored_errors( ).

    IF lt_ignored_errors IS NOT INITIAL.
      lo_element = io_document->create_simple_element( name   = 'ignoredErrors'
                                                       parent = io_document ).


      LOOP AT lt_ignored_errors ASSIGNING <ls_ignored_errors>.

        lo_element2 = io_document->create_simple_element( name   = 'ignoredError'
                                                          parent = io_document ).

        lo_element2->set_attribute_ns( name  = 'sqref'
                                       value = <ls_ignored_errors>-cell_coords ).

        IF <ls_ignored_errors>-eval_error = abap_true.
          lo_element2->set_attribute_ns( name  = 'evalError'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-two_digit_text_year = abap_true.
          lo_element2->set_attribute_ns( name  = 'twoDigitTextYear'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-number_stored_as_text = abap_true.
          lo_element2->set_attribute_ns( name  = 'numberStoredAsText'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'formula'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-formula_range = abap_true.
          lo_element2->set_attribute_ns( name  = 'formulaRange'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-unlocked_formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'unlockedFormula'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-empty_cell_reference = abap_true.
          lo_element2->set_attribute_ns( name  = 'emptyCellReference'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-list_data_validation = abap_true.
          lo_element2->set_attribute_ns( name  = 'listDataValidation'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-calculated_column = abap_true.
          lo_element2->set_attribute_ns( name  = 'calculatedColumn'
                                         value = '1' ).
        ENDIF.

        lo_element->append_child( lo_element2 ).

      ENDLOOP.

      io_element_root->append_child( lo_element ).

    ENDIF.

  ENDMETHOD.