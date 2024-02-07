  METHOD resolve_referenced_formulae.
    TYPES: BEGIN OF ty_referenced_cells,
             sheet    TYPE REF TO zcl_excel_worksheet,
             si       TYPE i,
             row_from TYPE i,
             row_to   TYPE i,
             col_from TYPE i,
             col_to   TYPE i,
             formula  TYPE string,
             ref_cell TYPE char10,
           END OF ty_referenced_cells.

    DATA: ls_ref_formula       LIKE LINE OF me->mt_ref_formulae,
          lts_referenced_cells TYPE SORTED TABLE OF ty_referenced_cells WITH NON-UNIQUE KEY sheet si row_from row_to col_from col_to,
          ls_referenced_cell   LIKE LINE OF lts_referenced_cells,
          lv_col_from          TYPE zexcel_cell_column_alpha,
          lv_col_to            TYPE zexcel_cell_column_alpha,
          lv_resulting_formula TYPE string,
          lv_current_cell      TYPE char10.


    me->mt_ref_formulae = me->mt_ref_formulae.

*--------------------------------------------------------------------*
* Get referenced Cells,  Build ranges for easy lookup
*--------------------------------------------------------------------*
    LOOP AT me->mt_ref_formulae INTO ls_ref_formula WHERE ref <> space.

      CLEAR ls_referenced_cell.
      ls_referenced_cell-sheet      = ls_ref_formula-sheet.
      ls_referenced_cell-si         = ls_ref_formula-si.
      ls_referenced_cell-formula    = ls_ref_formula-formula.

      TRY.
          zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = ls_ref_formula-ref
                                                        IMPORTING e_column_start = lv_col_from
                                                                  e_column_end   = lv_col_to
                                                                  e_row_start    = ls_referenced_cell-row_from
                                                                  e_row_end      = ls_referenced_cell-row_to  ).
          ls_referenced_cell-col_from = zcl_excel_common=>convert_column2int( lv_col_from ).
          ls_referenced_cell-col_to   = zcl_excel_common=>convert_column2int( lv_col_to ).


          CLEAR ls_referenced_cell-ref_cell.
          TRY.
              ls_referenced_cell-ref_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
              ls_referenced_cell-ref_cell+3  = ls_ref_formula-row.
              CONDENSE ls_referenced_cell-ref_cell NO-GAPS.
            CATCH zcx_excel.
          ENDTRY.

          INSERT ls_referenced_cell INTO TABLE lts_referenced_cells.
        CATCH zcx_excel.
      ENDTRY.

    ENDLOOP.

*  break x0009004.
*--------------------------------------------------------------------*
* For each referencing cell determine the referenced cell
* and resolve the formula
*--------------------------------------------------------------------*
    LOOP AT me->mt_ref_formulae INTO ls_ref_formula WHERE ref = space.


      CLEAR lv_current_cell.
      TRY.
          lv_current_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
          lv_current_cell+3  = ls_ref_formula-row.
          CONDENSE lv_current_cell NO-GAPS.
        CATCH zcx_excel.
      ENDTRY.

      LOOP AT lts_referenced_cells INTO ls_referenced_cell WHERE sheet     = ls_ref_formula-sheet
                                                             AND si        = ls_ref_formula-si
                                                             AND row_from <= ls_ref_formula-row
                                                             AND row_to   >= ls_ref_formula-row
                                                             AND col_from <= ls_ref_formula-column
                                                             AND col_to   >= ls_ref_formula-column.

        TRY.

            lv_resulting_formula = zcl_excel_common=>determine_resulting_formula( iv_reference_cell     = ls_referenced_cell-ref_cell
                                                                                  iv_reference_formula  = ls_referenced_cell-formula
                                                                                  iv_current_cell       = lv_current_cell ).

            ls_referenced_cell-sheet->set_cell_formula( ip_column   = ls_ref_formula-column
                                                        ip_row      = ls_ref_formula-row
                                                        ip_formula  = lv_resulting_formula ).
          CATCH zcx_excel.
        ENDTRY.
        EXIT.

      ENDLOOP.

    ENDLOOP.
  ENDMETHOD.